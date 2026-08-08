using System;
using System.Globalization;

namespace SteamRouteTool.Models
{
    /// <summary>A Steam application the tool can load a relay config for.</summary>
    public sealed class SteamGame
    {
        public SteamGame(int appId, string name)
        {
            AppId = appId;
            Name = name;
        }

        public int AppId { get; }

        public string Name { get; }

        /// <summary>Rendered directly by the app id combo box.</summary>
        public override string ToString()
        {
            return Name + " (" + AppId.ToString(CultureInfo.InvariantCulture) + ")";
        }
    }

    /// <summary>
    /// Convenience list for the app id prompt. It is not a whitelist: the Steam Web API
    /// serves a relay config for any valid app id, so anything can be typed in instead.
    /// </summary>
    public static class KnownGames
    {
        public static readonly SteamGame[] All =
        {
            new SteamGame(440, "Team Fortress 2"),
            new SteamGame(730, "Counter-Strike 2"),
            new SteamGame(570, "Dota 2"),
            new SteamGame(1422450, "Deadlock"),
            new SteamGame(252490, "Rust")
        };

        /// <summary>Returns the friendly name for an app id, or null when it is not in the list.</summary>
        public static string NameFor(int appId)
        {
            foreach (SteamGame game in All)
            {
                if (game.AppId == appId) return game.Name;
            }

            return null;
        }

        /// <summary>
        /// Accepts either a bare app id ("440") or a combo box entry
        /// ("Team Fortress 2 (440)").
        /// </summary>
        public static bool TryParseAppId(string text, out int appId)
        {
            appId = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            string candidate = text.Trim();

            int open = candidate.LastIndexOf('(');
            int close = candidate.LastIndexOf(')');
            if (open >= 0 && close > open)
            {
                candidate = candidate.Substring(open + 1, close - open - 1).Trim();
            }

            int parsed;
            if (!int.TryParse(candidate, NumberStyles.None, CultureInfo.InvariantCulture, out parsed)) return false;
            if (parsed <= 0) return false;

            appId = parsed;
            return true;
        }
    }
}
