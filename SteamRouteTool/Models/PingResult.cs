using System;
using System.Globalization;

namespace SteamRouteTool.Models
{
    /// <summary>Outcome of a single ICMP echo attempt.</summary>
    public struct PingResult
    {
        public static readonly PingResult Unreachable = new PingResult(false, -1);

        private PingResult(bool success, long roundtripMs)
        {
            Success = success;
            RoundtripMs = roundtripMs;
        }

        public bool Success { get; }

        /// <summary>Round trip in milliseconds, or -1 when the host did not reply.</summary>
        public long RoundtripMs { get; }

        public static PingResult FromReply(long roundtripMs)
        {
            return new PingResult(true, Math.Max(0, roundtripMs));
        }

        public override string ToString()
        {
            return Success ? RoundtripMs.ToString(CultureInfo.CurrentCulture) : "-";
        }
    }
}
