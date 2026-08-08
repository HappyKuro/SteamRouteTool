using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using SteamRouteTool.Models;

namespace SteamRouteTool.Services
{
    /// <summary>Raised when the SDR config cannot be retrieved or understood.</summary>
    [Serializable]
    public class SdrConfigException : Exception
    {
        public SdrConfigException(string message) : base(message) { }

        public SdrConfigException(string message, Exception inner) : base(message, inner) { }

        protected SdrConfigException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Downloads and parses Valve's Steam Datagram Relay network configuration
    /// (<c>ISteamApps/GetSDRConfig</c>) for a given app id.
    /// </summary>
    public sealed class SdrConfigClient
    {
        private const string ConfigUrlFormat =
            "https://api.steampowered.com/ISteamApps/GetSDRConfig/v1?appid={0}";

        private static readonly HttpClient Http = CreateClient();

        private readonly int _appId;

        public SdrConfigClient(int appId)
        {
            if (appId <= 0) throw new ArgumentOutOfRangeException("appId");
            _appId = appId;
        }

        private static HttpClient CreateClient()
        {
            // .NET Framework 4.7.2 honours the OS default, which on older machines still
            // excludes TLS 1.2 and makes the request fail with an opaque connection reset.
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch (NotSupportedException)
            {
                // Older platform without TLS 1.2; the request may still succeed.
            }

            var handler = new HttpClientHandler();
            if (handler.SupportsAutomaticDecompression)
            {
                handler.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            }

            var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(20) };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("SteamRouteTool/2.0");
            return client;
        }

        public async Task<IList<PointOfPresence>> GetPointsOfPresenceAsync(CancellationToken cancellationToken)
        {
            string url = string.Format(CultureInfo.InvariantCulture, ConfigUrlFormat, _appId);

            string body;
            try
            {
                using (HttpResponseMessage response = await Http.GetAsync(url, cancellationToken).ConfigureAwait(false))
                {
                    body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new SdrConfigException(DescribeFailure(response, body, _appId));
                    }
                }
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (SdrConfigException)
            {
                throw;
            }
            catch (Exception ex)
            {
                // A client timeout also surfaces as a cancellation, but the caller did not
                // ask for one, so it has to be reported rather than silently ignored.
                throw new SdrConfigException("Could not reach the Steam Web API. Check your connection.", ex);
            }

            return Parse(body);
        }

        /// <summary>
        /// Turns a failed response into something the user can act on. An app id Steam does
        /// not recognise comes back as HTTP 500 with
        /// <c>{"success":false,"message":"Failed to get appinfo"}</c>.
        /// </summary>
        internal static string DescribeFailure(HttpResponseMessage response, string body, int appId)
        {
            string reported = ReadApiMessage(body);
            if (reported != null)
            {
                return string.Format(CultureInfo.CurrentCulture,
                    "Steam does not have a relay config for app id {0} ({1}). Check the app ID and try again.",
                    appId, reported);
            }

            return string.Format(CultureInfo.CurrentCulture,
                "Steam returned {0} ({1}) for app id {2}.",
                (int)response.StatusCode, response.ReasonPhrase, appId);
        }

        private static string ReadApiMessage(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return null;

            try
            {
                var message = (string)JObject.Parse(body)["message"];
                return string.IsNullOrWhiteSpace(message) ? null : message;
            }
            catch (Exception)
            {
                return null; // Body was not the JSON error shape; fall back to the status line.
            }
        }

        /// <summary>Parses a GetSDRConfig response body. Internal so it can be exercised without network access.</summary>
        internal static IList<PointOfPresence> Parse(string json)
        {
            JObject root;
            try
            {
                root = JObject.Parse(json);
            }
            catch (Exception ex)
            {
                throw new SdrConfigException("The Steam Web API returned a response that could not be parsed.", ex);
            }

            var pops = root["pops"] as JObject;
            if (pops == null)
            {
                throw new SdrConfigException("The Steam Web API response did not contain any points of presence.");
            }

            var result = new List<PointOfPresence>();
            foreach (JProperty property in pops.Properties())
            {
                PointOfPresence pop = ParsePop(property.Name, property.Value as JObject);
                if (pop != null) result.Add(pop);
            }

            if (result.Count == 0)
            {
                throw new SdrConfigException("The Steam Web API response contained no relays to display.");
            }

            result.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.CurrentCultureIgnoreCase));
            return result;
        }

        private static PointOfPresence ParsePop(string code, JObject value)
        {
            if (value == null) return null;

            var relayArray = value["relays"] as JArray;
            if (relayArray == null) return null;

            var description = (string)value["desc"];
            if (IsCloudTest(code) || IsCloudTest(description)) return null;

            var relays = new List<Relay>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (JObject relayToken in relayArray.OfType<JObject>())
            {
                var ipv4 = (string)relayToken["ipv4"];
                if (string.IsNullOrWhiteSpace(ipv4)) continue;

                ipv4 = ipv4.Trim();
                if (!seen.Add(ipv4)) continue; // Valve occasionally repeats an address within a PoP.

                relays.Add(new Relay(ipv4, ParsePortRange(relayToken["port_range"])));
            }

            if (relays.Count == 0) return null;

            return new PointOfPresence(
                code,
                description,
                (int?)value["partners"] ?? 0,
                (int?)value["tier"] ?? 0,
                relays);
        }

        private static bool IsCloudTest(string text)
        {
            return text != null && text.IndexOf("cloud-test", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// "port_range" is a two element array such as [27015, 27060]; ranges differ per PoP.
        /// </summary>
        private static PortRange ParsePortRange(JToken token)
        {
            var array = token as JArray;
            if (array != null && array.Count >= 2)
            {
                int? low = (int?)array[0];
                int? high = (int?)array[1];
                if (low.HasValue && high.HasValue &&
                    low.Value >= 0 && high.Value >= low.Value && high.Value <= 65535)
                {
                    return new PortRange(low.Value, high.Value);
                }
            }

            return PortRange.SteamDefault;
        }
    }
}
