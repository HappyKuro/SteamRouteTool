using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SteamRouteTool.Models
{
    /// <summary>
    /// A Valve point of presence (data centre) and the relays it exposes.
    /// </summary>
    public sealed class PointOfPresence
    {
        public PointOfPresence(string code, string description, int partners, int tier, IList<Relay> relays)
        {
            if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Code is required.", "code");
            if (relays == null) throw new ArgumentNullException("relays");

            Code = code;
            Description = description;
            Partners = partners;
            Tier = tier;
            Relays = new ReadOnlyCollection<Relay>(new List<Relay>(relays));
        }

        /// <summary>Short Valve code, e.g. "ams". Used to name the firewall rules.</summary>
        public string Code { get; }

        /// <summary>Human readable location, e.g. "Amsterdam (Netherlands)". May be null.</summary>
        public string Description { get; }

        /// <summary>Bit mask of the Valve partner networks that serve this PoP.</summary>
        public int Partners { get; }

        public int Tier { get; }

        public ReadOnlyCollection<Relay> Relays { get; }

        public string DisplayName
        {
            get { return string.IsNullOrWhiteSpace(Description) ? Code : Description; }
        }

        public override string ToString()
        {
            return Code + " (" + Relays.Count + " relays)";
        }
    }
}
