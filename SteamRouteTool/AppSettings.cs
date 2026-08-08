using System;
using System.Configuration;
using System.Globalization;

namespace SteamRouteTool
{
    /// <summary>
    /// Tunables read from App.config, each with a sensible default so the tool still runs
    /// with an empty or damaged configuration file.
    /// </summary>
    internal static class AppSettings
    {
        /// <summary>Team Fortress 2. Any app that uses GetSDRConfig works here.</summary>
        public const int DefaultAppId = 440;

        /// <summary>App id seeded from App.config, used the first time the tool runs.</summary>
        public static int AppId
        {
            get { return ReadInt32("appId", DefaultAppId, 1, int.MaxValue); }
        }

        /// <summary>
        /// What the app id prompt should offer: whatever was chosen last, falling back to
        /// the App.config value on a first run.
        /// </summary>
        public static int InitialAppId
        {
            get
            {
                try
                {
                    int saved = Properties.Settings.Default.LastAppId;
                    if (saved > 0) return saved;
                }
                catch (ConfigurationException)
                {
                    // Damaged user.config; fall back to the shipped default.
                }

                return AppId;
            }
        }

        /// <summary>Remembers the chosen app id so the next launch offers it first.</summary>
        public static void SaveLastAppId(int appId)
        {
            if (appId <= 0) return;

            try
            {
                Properties.Settings.Default.LastAppId = appId;
                Properties.Settings.Default.Save();
            }
            catch (ConfigurationException)
            {
                // Remembering the choice is a convenience, never a reason to fail.
            }
        }

        public static int PingTimeoutMs
        {
            get { return ReadInt32("pingTimeoutMs", 1000, 100, 10000); }
        }

        public static int MaxConcurrentPings
        {
            get { return ReadInt32("maxConcurrentPings", 16, 1, 64); }
        }

        /// <summary>Latency at or below this is shown as good.</summary>
        public static int GoodPingMs
        {
            get { return ReadInt32("goodPingMs", 50, 1, 10000); }
        }

        /// <summary>Latency at or below this is shown as acceptable; anything higher is poor.</summary>
        public static int FairPingMs
        {
            get { return ReadInt32("fairPingMs", 100, 1, 10000); }
        }

        private static int ReadInt32(string key, int fallback, int min, int max)
        {
            string raw;
            try
            {
                raw = ConfigurationManager.AppSettings[key];
            }
            catch (ConfigurationErrorsException)
            {
                return fallback;
            }

            int value;
            if (string.IsNullOrWhiteSpace(raw) ||
                !int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                return fallback;
            }

            return Math.Min(Math.Max(value, min), max);
        }
    }
}
