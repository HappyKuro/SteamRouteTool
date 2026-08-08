using System;
using System.Collections.Generic;
using System.Globalization;

namespace SteamRouteTool.Models
{
    /// <summary>
    /// An inclusive TCP/UDP port range, as advertised by a relay in the SDR config.
    /// </summary>
    public struct PortRange : IEquatable<PortRange>
    {
        /// <summary>
        /// Range used when a relay does not advertise one. Wide enough to cover every
        /// range Valve currently publishes.
        /// </summary>
        public static readonly PortRange SteamDefault = new PortRange(27015, 27202);

        public PortRange(int low, int high)
        {
            if (low < 0 || low > 65535)
                throw new ArgumentOutOfRangeException("low");
            if (high < low || high > 65535)
                throw new ArgumentOutOfRangeException("high");

            Low = low;
            High = high;
        }

        public int Low { get; }

        public int High { get; }

        public override string ToString()
        {
            return Low == High
                ? Low.ToString(CultureInfo.InvariantCulture)
                : Low.ToString(CultureInfo.InvariantCulture) + "-" + High.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Collapses overlapping and adjacent ranges into the smallest equivalent set.
        /// </summary>
        public static List<PortRange> Merge(IEnumerable<PortRange> ranges)
        {
            if (ranges == null) throw new ArgumentNullException("ranges");

            var ordered = new List<PortRange>(ranges);
            ordered.Sort((a, b) => a.Low != b.Low ? a.Low.CompareTo(b.Low) : a.High.CompareTo(b.High));

            var merged = new List<PortRange>();
            foreach (PortRange range in ordered)
            {
                if (merged.Count == 0)
                {
                    merged.Add(range);
                    continue;
                }

                PortRange last = merged[merged.Count - 1];
                if (range.Low > last.High + 1)
                {
                    merged.Add(range);
                }
                else if (range.High > last.High)
                {
                    merged[merged.Count - 1] = new PortRange(last.Low, range.High);
                }
            }

            return merged;
        }

        /// <summary>
        /// Formats ranges the way the Windows Firewall <c>RemotePorts</c> property expects
        /// (a comma separated list of ports and "low-high" ranges).
        /// </summary>
        public static string Format(IEnumerable<PortRange> ranges)
        {
            List<PortRange> merged = Merge(ranges);
            var parts = new string[merged.Count];
            for (int i = 0; i < merged.Count; i++) parts[i] = merged[i].ToString();
            return string.Join(",", parts);
        }

        public bool Equals(PortRange other)
        {
            return Low == other.Low && High == other.High;
        }

        public override bool Equals(object obj)
        {
            return obj is PortRange && Equals((PortRange)obj);
        }

        public override int GetHashCode()
        {
            return (Low * 397) ^ High;
        }
    }
}
