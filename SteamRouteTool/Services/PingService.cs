using System;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using SteamRouteTool.Models;

namespace SteamRouteTool.Services
{
    /// <summary>Asynchronous ICMP echo helper.</summary>
    public static class PingService
    {
        /// <summary>
        /// Pings <paramref name="host"/> without blocking a thread pool thread. Never throws
        /// for an unreachable host; failures come back as <see cref="PingResult.Unreachable"/>.
        /// </summary>
        public static async Task<PingResult> SendAsync(string host, int timeoutMs, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(host)) return PingResult.Unreachable;
            cancellationToken.ThrowIfCancellationRequested();

            using (var ping = new Ping())
            using (cancellationToken.Register(() => TryCancel(ping)))
            {
                try
                {
                    PingReply reply = await ping.SendPingAsync(host, timeoutMs).ConfigureAwait(false);

                    // A sub-millisecond reply legitimately reports 0ms, so only Status decides.
                    return reply != null && reply.Status == IPStatus.Success
                        ? PingResult.FromReply(reply.RoundtripTime)
                        : PingResult.Unreachable;
                }
                catch (PingException)
                {
                    return PingResult.Unreachable;
                }
                catch (InvalidOperationException)
                {
                    // Raised when the send is cancelled from under us.
                    return PingResult.Unreachable;
                }
            }
        }

        private static void TryCancel(Ping ping)
        {
            try
            {
                ping.SendAsyncCancel();
            }
            catch (Exception)
            {
                // The ping already finished; nothing to cancel.
            }
        }
    }
}
