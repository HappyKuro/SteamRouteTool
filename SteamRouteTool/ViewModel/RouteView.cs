using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using SteamRouteTool.Models;

namespace SteamRouteTool.ViewModel
{
    /// <summary>How the location list is ordered.</summary>
    public enum RouteSort
    {
        Location,
        Ping
    }

    /// <summary>What a grid row represents.</summary>
    public enum RouteLineKind
    {
        /// <summary>A location: expandable, and its checkbox covers every relay it holds.</summary>
        Location,

        /// <summary>A single relay underneath an expanded location.</summary>
        Relay
    }

    /// <summary>One rendered grid row.</summary>
    public sealed class RouteLine
    {
        internal RouteLine(PopGroup group, RouteRow row)
        {
            Group = group;
            Row = row;
        }

        public PopGroup Group { get; }

        /// <summary>The relay this line shows, or null on a location line.</summary>
        public RouteRow Row { get; }

        public RouteLineKind Kind
        {
            get { return Row == null ? RouteLineKind.Location : RouteLineKind.Relay; }
        }
    }

    /// <summary>
    /// The model behind the grid: locations, the relays under them, and the expand, block and
    /// latency state the UI renders. <see cref="BuildLines"/> turns all of that into the exact
    /// list of rows to draw, so the form never has to reason about row indexes.
    /// </summary>
    public sealed class RouteView
    {
        public static readonly RouteView Empty = new RouteView(new PointOfPresence[0]);

        private readonly List<RouteRow> _rows = new List<RouteRow>();
        private readonly List<PopGroup> _groups = new List<PopGroup>();

        public RouteView(IEnumerable<PointOfPresence> pops)
        {
            if (pops == null) throw new ArgumentNullException("pops");

            foreach (PointOfPresence pop in pops)
            {
                var groupRows = new List<RouteRow>(pop.Relays.Count);
                var group = new PopGroup(pop, groupRows);

                for (int i = 0; i < pop.Relays.Count; i++)
                {
                    var row = new RouteRow(group, pop.Relays[i], i);
                    groupRows.Add(row);
                    _rows.Add(row);
                }

                _groups.Add(group);
            }

            Rows = new ReadOnlyCollection<RouteRow>(_rows);
            Groups = new ReadOnlyCollection<PopGroup>(_groups);
        }

        public ReadOnlyCollection<RouteRow> Rows { get; }

        public ReadOnlyCollection<PopGroup> Groups { get; }

        public int RelayCount
        {
            get { return _rows.Count; }
        }

        /// <summary>
        /// Produces the rows to draw: every matching location, each followed by its relays when
        /// it is expanded.
        /// </summary>
        /// <param name="filter">Matched against location name, location code and relay address. May be null.</param>
        /// <param name="sort">Order of the locations. Relays always keep their published order.</param>
        /// <param name="descending">Reverses the location order.</param>
        public List<RouteLine> BuildLines(string filter, RouteSort sort, bool descending)
        {
            var groups = new List<PopGroup>();
            foreach (PopGroup group in _groups)
            {
                if (Matches(group, filter)) groups.Add(group);
            }

            groups.Sort((a, b) => Compare(a, b, sort));
            if (descending) groups.Reverse();

            var lines = new List<RouteLine>();
            foreach (PopGroup group in groups)
            {
                lines.Add(new RouteLine(group, null));
                if (!group.IsExpanded) continue;

                foreach (RouteRow row in group.Rows)
                {
                    lines.Add(new RouteLine(group, row));
                }
            }

            return lines;
        }

        private static int Compare(PopGroup a, PopGroup b, RouteSort sort)
        {
            if (sort == RouteSort.Ping)
            {
                // Locations that have not answered yet sort last, so the list stays useful
                // while a sweep is still running.
                long left = a.BestPing.HasValue ? a.BestPing.Value : long.MaxValue;
                long right = b.BestPing.HasValue ? b.BestPing.Value : long.MaxValue;
                if (left != right) return left.CompareTo(right);
            }

            return string.Compare(a.Pop.DisplayName, b.Pop.DisplayName, StringComparison.CurrentCultureIgnoreCase);
        }

        private static bool Matches(PopGroup group, string filter)
        {
            if (string.IsNullOrWhiteSpace(filter)) return true;

            string needle = filter.Trim();
            if (Contains(group.Pop.DisplayName, needle)) return true;
            if (Contains(group.Pop.Code, needle)) return true;

            foreach (RouteRow row in group.Rows)
            {
                if (Contains(row.Relay.Ipv4, needle)) return true;
            }

            return false;
        }

        private static bool Contains(string haystack, string needle)
        {
            return haystack != null && haystack.IndexOf(needle, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }
    }

    /// <summary>A location together with its relays, expand state and aggregate latency.</summary>
    public sealed class PopGroup
    {
        private readonly List<RouteRow> _rows;

        internal PopGroup(PointOfPresence pop, List<RouteRow> rows)
        {
            Pop = pop;
            _rows = rows;
            Rows = new ReadOnlyCollection<RouteRow>(rows);
        }

        public PointOfPresence Pop { get; }

        public ReadOnlyCollection<RouteRow> Rows { get; }

        public bool IsExpanded { get; set; }

        /// <summary>Whether this location holds more than the one relay shown when collapsed.</summary>
        public bool CanExpand
        {
            get { return _rows.Count > 1; }
        }

        public bool AllBlocked
        {
            get
            {
                foreach (RouteRow row in _rows)
                {
                    if (!row.IsBlocked) return false;
                }
                return _rows.Count > 0;
            }
        }

        public bool AnyBlocked
        {
            get
            {
                foreach (RouteRow row in _rows)
                {
                    if (row.IsBlocked) return true;
                }
                return false;
            }
        }

        public bool IsPartiallyBlocked
        {
            get { return AnyBlocked && !AllBlocked; }
        }

        /// <summary>Lowest latency any relay here reported, or null when none has answered.</summary>
        public long? BestPing
        {
            get
            {
                long? best = null;
                foreach (RouteRow row in _rows)
                {
                    if (!row.LastPing.HasValue) continue;

                    PingResult ping = row.LastPing.Value;
                    if (!ping.Success) continue;
                    if (!best.HasValue || ping.RoundtripMs < best.Value) best = ping.RoundtripMs;
                }

                return best;
            }
        }

        /// <summary>True while any relay here is waiting on a reply.</summary>
        public bool IsPinging
        {
            get
            {
                foreach (RouteRow row in _rows)
                {
                    if (row.IsPinging) return true;
                }
                return false;
            }
        }

        /// <summary>Every relay here has been pinged and none answered.</summary>
        public bool AllUnreachable
        {
            get
            {
                foreach (RouteRow row in _rows)
                {
                    if (!row.LastPing.HasValue || row.LastPing.Value.Success) return false;
                }
                return _rows.Count > 0;
            }
        }

        public List<Relay> BlockedRelays()
        {
            var blocked = new List<Relay>();
            foreach (RouteRow row in _rows)
            {
                if (row.IsBlocked) blocked.Add(row.Relay);
            }
            return blocked;
        }

        public void SetAllBlocked(bool blocked)
        {
            foreach (RouteRow row in _rows) row.IsBlocked = blocked;
        }
    }

    /// <summary>One relay and its current state.</summary>
    public sealed class RouteRow
    {
        internal RouteRow(PopGroup group, Relay relay, int indexInPop)
        {
            Group = group;
            Relay = relay;
            IndexInPop = indexInPop;
        }

        public PopGroup Group { get; }

        public Relay Relay { get; }

        /// <summary>Zero based position of this relay within its location.</summary>
        public int IndexInPop { get; }

        public bool IsBlocked { get; set; }

        /// <summary>Latest ping outcome, or null if this relay has not been pinged yet.</summary>
        public PingResult? LastPing { get; set; }

        /// <summary>True while a reply is outstanding.</summary>
        public bool IsPinging { get; set; }
    }
}
