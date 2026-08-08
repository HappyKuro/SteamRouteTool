using System;

namespace SteamRouteTool.Models
{
    /// <summary>
    /// A single Steam Datagram Relay endpoint inside a point of presence.
    /// </summary>
    public sealed class Relay
    {
        public Relay(string ipv4, PortRange ports)
        {
            if (string.IsNullOrWhiteSpace(ipv4)) throw new ArgumentException("Address is required.", "ipv4");

            Ipv4 = ipv4;
            Ports = ports;
        }

        public string Ipv4 { get; }

        public PortRange Ports { get; }

        public override string ToString()
        {
            return Ipv4 + ":" + Ports;
        }
    }
}
