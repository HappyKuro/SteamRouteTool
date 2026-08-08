using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using NetFwTypeLib;
using SteamRouteTool.Models;

namespace SteamRouteTool.Services
{
    /// <summary>Raised when a Windows Firewall operation cannot be completed.</summary>
    [Serializable]
    public class FirewallException : Exception
    {
        public FirewallException(string message) : base(message) { }

        public FirewallException(string message, Exception inner) : base(message, inner) { }

        protected FirewallException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }

    /// <summary>The set of relays that should be blocked for one point of presence.</summary>
    public sealed class BlockRequest
    {
        public BlockRequest(string popCode, IList<Relay> relays)
        {
            if (string.IsNullOrWhiteSpace(popCode)) throw new ArgumentException("PoP code is required.", "popCode");

            PopCode = popCode;
            Relays = new List<Relay>(relays ?? new Relay[0]);
        }

        public string PopCode { get; }

        /// <summary>Relays to block. Empty means "remove this PoP's rules".</summary>
        public List<Relay> Relays { get; }
    }

    /// <summary>
    /// Creates and removes the outbound block rules this tool owns, through the
    /// Windows Firewall COM API. Every rule it writes is named
    /// <c>SteamRouteTool-{PROTOCOL}-{popCode}</c>, which is also how it finds them again.
    /// </summary>
    /// <remarks>
    /// All public members serialise on a single lock: the firewall store does not react
    /// well to concurrent writers, and the UI can easily request several updates at once.
    /// </remarks>
    public sealed class FirewallService
    {
        public const string RulePrefix = "SteamRouteTool-";

        /// <summary>Rules written by the older TF2RoutingTool, cleared on start-up.</summary>
        public const string LegacyRulePrefix = "TF2RoutingTool-";

        private const string RuleGroup = "SteamRouteTool";
        private const int ProtocolIcmpV4 = 1;
        private const int ProtocolTcp = 6;
        private const int ProtocolUdp = 17;

        private readonly object _gate = new object();

        /// <summary>
        /// Rewrites the rules for each requested PoP: existing rules are dropped and, when
        /// the request lists relays, fresh TCP/UDP/ICMP block rules are created.
        /// </summary>
        public void ApplyBlocks(IEnumerable<BlockRequest> requests)
        {
            if (requests == null) throw new ArgumentNullException("requests");

            List<BlockRequest> batch = requests.ToList();
            if (batch.Count == 0) return;

            lock (_gate)
            {
                INetFwPolicy2 policy = CreatePolicy();
                HashSet<string> existing = ReadRuleNames(policy);

                foreach (BlockRequest request in batch)
                {
                    foreach (string name in RuleNamesFor(request.PopCode))
                    {
                        RemoveRule(policy, existing, name);
                    }

                    if (request.Relays.Count == 0) continue;

                    string addresses = string.Join(",", request.Relays.Select(r => r.Ipv4));
                    string ports = PortRange.Format(request.Relays.Select(r => r.Ports));

                    AddRule(policy, RuleName("UDP", request.PopCode), ProtocolUdp, addresses, ports, request.PopCode);
                    AddRule(policy, RuleName("TCP", request.PopCode), ProtocolTcp, addresses, ports, request.PopCode);
                    AddRule(policy, RuleName("ICMP", request.PopCode), ProtocolIcmpV4, addresses, null, request.PopCode);
                }
            }
        }

        /// <summary>Removes every rule whose name starts with <paramref name="prefix"/>. Returns the number removed.</summary>
        public int RemoveRulesWithPrefix(string prefix)
        {
            if (string.IsNullOrEmpty(prefix)) throw new ArgumentException("Prefix is required.", "prefix");

            lock (_gate)
            {
                INetFwPolicy2 policy = CreatePolicy();
                List<string> targets = ReadRuleNames(policy)
                    .Where(name => name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                int removed = 0;
                foreach (string name in targets)
                {
                    if (RemoveRule(policy, null, name)) removed++;
                }

                return removed;
            }
        }

        /// <summary>
        /// Reads back which addresses are currently blocked, keyed by PoP code, so the UI can
        /// restore its state after a restart.
        /// </summary>
        public IDictionary<string, HashSet<string>> ReadBlockedAddresses()
        {
            var blocked = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            lock (_gate)
            {
                INetFwPolicy2 policy = CreatePolicy();

                foreach (INetFwRule rule in EnumerateRules(policy))
                {
                    string popCode = PopCodeFromRuleName(rule.Name);
                    if (popCode == null) continue;

                    string remoteAddresses;
                    try
                    {
                        remoteAddresses = rule.RemoteAddresses;
                    }
                    catch (COMException)
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(remoteAddresses) || remoteAddresses == "*") continue;

                    HashSet<string> addresses;
                    if (!blocked.TryGetValue(popCode, out addresses))
                    {
                        addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        blocked[popCode] = addresses;
                    }

                    // Windows reports addresses as "1.2.3.4/255.255.255.255,5.6.7.8/255.255.255.255".
                    foreach (string entry in remoteAddresses.Split(','))
                    {
                        string address = entry.Split('/')[0].Trim();
                        if (address.Length > 0) addresses.Add(address);
                    }
                }
            }

            return blocked;
        }

        private static IEnumerable<string> RuleNamesFor(string popCode)
        {
            yield return RuleName("UDP", popCode);
            yield return RuleName("TCP", popCode);
            yield return RuleName("ICMP", popCode);
        }

        private static string RuleName(string protocol, string popCode)
        {
            return RulePrefix + protocol + "-" + popCode;
        }

        /// <summary>Extracts the PoP code from one of our rule names, or null if the rule is not ours.</summary>
        private static string PopCodeFromRuleName(string ruleName)
        {
            if (ruleName == null) return null;
            if (!ruleName.StartsWith(RulePrefix, StringComparison.OrdinalIgnoreCase)) return null;

            // Skip the protocol segment; whatever follows is the PoP code, dashes included.
            int separator = ruleName.IndexOf('-', RulePrefix.Length);
            if (separator < 0 || separator + 1 >= ruleName.Length) return null;

            return ruleName.Substring(separator + 1);
        }

        private static INetFwPolicy2 CreatePolicy()
        {
            try
            {
                Type policyType = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                if (policyType == null)
                {
                    throw new FirewallException("The Windows Firewall COM interface is not registered on this machine.");
                }

                return (INetFwPolicy2)Activator.CreateInstance(policyType);
            }
            catch (FirewallException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new FirewallException("Could not open the Windows Firewall. Is the Windows Defender Firewall service running?", ex);
            }
        }

        private static INetFwRule CreateRuleObject()
        {
            Type ruleType = Type.GetTypeFromProgID("HNetCfg.FWRule");
            if (ruleType == null)
            {
                throw new FirewallException("The Windows Firewall COM interface is not registered on this machine.");
            }

            return (INetFwRule)Activator.CreateInstance(ruleType);
        }

        private static void AddRule(INetFwPolicy2 policy, string name, int protocol, string remoteAddresses, string remotePorts, string popCode)
        {
            try
            {
                INetFwRule rule = CreateRuleObject();
                rule.Name = name;
                rule.Description = string.Format(CultureInfo.InvariantCulture,
                    "Blocks outbound traffic to the {0} Steam relay. Created by SteamRouteTool.", popCode);
                rule.Grouping = RuleGroup;
                rule.Enabled = true;
                rule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
                rule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
                rule.Profiles = (int)NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_ALL;
                rule.Protocol = protocol;
                rule.RemoteAddresses = remoteAddresses;

                // Ports are only valid on TCP and UDP rules; setting them on ICMP throws.
                if (!string.IsNullOrEmpty(remotePorts) && (protocol == ProtocolTcp || protocol == ProtocolUdp))
                {
                    rule.RemotePorts = remotePorts;
                }

                policy.Rules.Add(rule);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new FirewallException("Administrator rights are required to change firewall rules.", ex);
            }
            catch (COMException ex)
            {
                throw new FirewallException(
                    string.Format(CultureInfo.CurrentCulture, "Windows rejected the firewall rule \"{0}\".", name), ex);
            }
        }

        /// <summary>
        /// Removes a rule by name. <paramref name="known"/>, when supplied, avoids the
        /// exception Windows raises for a name that does not exist.
        /// </summary>
        private static bool RemoveRule(INetFwPolicy2 policy, HashSet<string> known, string name)
        {
            if (known != null && !known.Contains(name)) return false;

            try
            {
                policy.Rules.Remove(name);
                if (known != null) known.Remove(name);
                return true;
            }
            catch (FileNotFoundException)
            {
                return false;
            }
            catch (COMException ex)
            {
                Debug.WriteLine("FirewallService: could not remove rule '" + name + "': " + ex.Message);
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new FirewallException("Administrator rights are required to change firewall rules.", ex);
            }
        }

        private static HashSet<string> ReadRuleNames(INetFwPolicy2 policy)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (INetFwRule rule in EnumerateRules(policy))
            {
                if (rule.Name != null) names.Add(rule.Name);
            }

            return names;
        }

        /// <summary>
        /// Snapshots the rule collection. Enumeration is done once per operation because it
        /// is by far the most expensive part of talking to the firewall.
        /// </summary>
        private static List<INetFwRule> EnumerateRules(INetFwPolicy2 policy)
        {
            var rules = new List<INetFwRule>();
            try
            {
                foreach (INetFwRule rule in policy.Rules)
                {
                    rules.Add(rule);
                }
            }
            catch (COMException ex)
            {
                // A single corrupt rule can abort enumeration; keep whatever was read.
                Debug.WriteLine("FirewallService: rule enumeration stopped early: " + ex.Message);
            }

            return rules;
        }
    }
}
