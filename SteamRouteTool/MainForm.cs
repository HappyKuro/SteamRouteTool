using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using SteamRouteTool.Models;
using SteamRouteTool.Services;
using SteamRouteTool.ViewModel;

namespace SteamRouteTool
{
    /// <summary>
    /// Lists the Steam Datagram Relay locations for the configured app and lets the user block
    /// them with outbound Windows Firewall rules.
    /// </summary>
    /// <remarks>
    /// Every member here runs on the UI thread. Slow work (firewall COM calls, ICMP echoes) is
    /// awaited rather than posted to a background thread that then reaches back into the grid,
    /// so no marshalling helper is needed and the grid is never touched off-thread.
    /// </remarks>
    public partial class MainForm : Form
    {
        private const int ColumnName = 0;
        private const int ColumnPing = 1;
        private const int ColumnBlocked = 2;

        /// <summary>Latency that fills the bar in the ping column.</summary>
        private const int PingBarScaleMs = 200;

        private const int EM_SETCUEBANNER = 0x1501;

        private static readonly Color TextColor = Color.FromArgb(32, 36, 42);
        private static readonly Color MutedColor = Color.FromArgb(122, 128, 138);
        private static readonly Color HoverColor = Color.FromArgb(242, 246, 252);
        private static readonly Color BlockedTint = Color.FromArgb(253, 241, 241);
        private static readonly Color TrackColor = Color.FromArgb(232, 234, 238);
        private static readonly Color GoodColor = Color.FromArgb(30, 126, 52);
        private static readonly Color FairColor = Color.FromArgb(197, 106, 10);
        private static readonly Color PoorColor = Color.FromArgb(190, 45, 45);
        private static readonly Color DeadColor = Color.FromArgb(150, 152, 158);

        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private readonly FirewallService _firewall = new FirewallService();

        private readonly int _pingTimeoutMs = AppSettings.PingTimeoutMs;
        private readonly int _maxConcurrentPings = AppSettings.MaxConcurrentPings;
        private readonly int _goodPingMs = AppSettings.GoodPingMs;
        private readonly int _fairPingMs = AppSettings.FairPingMs;

        private readonly List<RouteLine> _lines = new List<RouteLine>();
        private readonly Dictionary<RouteRow, int> _relayRowIndex = new Dictionary<RouteRow, int>();
        private readonly Dictionary<PopGroup, int> _groupRowIndex = new Dictionary<PopGroup, int>();

        private RouteView _view = RouteView.Empty;
        private RouteSort _sort = RouteSort.Location;
        private bool _sortDescending;
        private CancellationTokenSource _sweep;
        private int _busyDepth;
        private int _appId;
        private int _hoverRowIndex = -1;
        private RouteLine _menuLine;
        private bool _routesLoaded;

        /// <summary>Guards against reacting to cell values that we wrote ourselves.</summary>
        private bool _updatingGrid;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        /// <param name="appId">Steam app whose relay config is loaded, chosen in the start-up prompt.</param>
        public MainForm(int appId)
        {
            if (appId <= 0) throw new ArgumentOutOfRangeException("appId");

            InitializeComponent();
            EnableDoubleBuffering(routeGrid);

            colPing.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            colPing.HeaderCell.Style.Padding = new Padding(0, 0, 12, 0);

            _appId = appId;
            UpdateTitle();
            UpdateHeaders();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            SendMessage(txtFilter.Handle, EM_SETCUEBANNER, (IntPtr)1, "Filter by location or IP address");
        }

        #region Start-up and loading

        private async void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                await StartUpAsync();
            }
            catch (OperationCanceledException)
            {
                // The form is closing.
            }
            catch (SdrConfigException ex)
            {
                ShowOverlay("Could not load the route list.");
                SetStatus("Could not load the route list.");
                ShowError("Could not load the route list.", ex);
            }
            catch (Exception ex)
            {
                ShowOverlay("Start-up failed.");
                SetStatus("Start-up failed.");
                ShowError("SteamRouteTool could not start.", ex);
            }
        }

        private async Task StartUpAsync()
        {
            CancellationToken token = _shutdown.Token;

            BeginBusy("Clearing rules left behind by TF2RoutingTool...");
            try
            {
                await Task.Run(() => _firewall.RemoveRulesWithPrefix(FirewallService.LegacyRulePrefix), token);
            }
            catch (FirewallException ex)
            {
                // Not fatal: the user may simply not have those rules, or not be elevated yet.
                Debug.WriteLine("Legacy rule clean-up skipped: " + ex.Message);
            }
            finally
            {
                EndBusy();
            }

            await LoadGameAsync(_appId);
        }

        /// <summary>
        /// Loads the relay config for an app and rebuilds the grid around it. Nothing on the
        /// form changes until the download succeeds, so a failed switch leaves the current
        /// game on screen.
        /// </summary>
        private async Task LoadGameAsync(int appId)
        {
            CancellationToken token = _shutdown.Token;
            var client = new SdrConfigClient(appId);

            ShowOverlay("Loading routes...");
            BeginBusy(string.Format(CultureInfo.CurrentCulture,
                "Downloading the route list for app id {0}...", appId));
            IList<PointOfPresence> pops;
            try
            {
                pops = await client.GetPointsOfPresenceAsync(token);
            }
            finally
            {
                EndBusy();
            }

            _appId = appId;
            AppSettings.SaveLastAppId(appId);
            UpdateTitle();

            _view = new RouteView(pops);
            _routesLoaded = true;
            txtFilter.Clear();
            RenderLines();

            btnPingRoutes.Enabled = true;
            btnClearRules.Enabled = true;

            BeginBusy("Reading existing firewall rules...");
            try
            {
                IDictionary<string, HashSet<string>> blocked =
                    await Task.Run(() => _firewall.ReadBlockedAddresses(), token);
                RestoreBlockedState(blocked);
            }
            catch (FirewallException ex)
            {
                Debug.WriteLine("Could not read existing rules: " + ex.Message);
            }
            finally
            {
                EndBusy();
            }

            SetStatus(DescribeSelection());
            await PingVisibleRoutesAsync();
        }

        /// <summary>Ticks the checkboxes for relays that are already blocked by an existing rule.</summary>
        private void RestoreBlockedState(IDictionary<string, HashSet<string>> blockedByPop)
        {
            foreach (PopGroup group in _view.Groups)
            {
                HashSet<string> addresses;
                if (!blockedByPop.TryGetValue(group.Pop.Code, out addresses)) continue;

                foreach (RouteRow row in group.Rows)
                {
                    row.IsBlocked = addresses.Contains(row.Relay.Ipv4);
                }
            }

            RenderLines();
        }

        private void UpdateTitle()
        {
            string name = KnownGames.NameFor(_appId);
            string game = name != null
                ? name + " (" + _appId.ToString(CultureInfo.CurrentCulture) + ")"
                : "app id " + _appId.ToString(CultureInfo.CurrentCulture);

            Text = "SteamRouteTool - " + game;
        }

        #endregion

        #region Rendering

        /// <summary>
        /// Rebuilds the grid from the view model, honouring the current filter, sort order and
        /// which locations are expanded.
        /// </summary>
        private void RenderLines()
        {
            List<RouteLine> lines = _view.BuildLines(txtFilter.Text, _sort, _sortDescending);

            _updatingGrid = true;
            routeGrid.SuspendLayout();
            try
            {
                routeGrid.CurrentCell = null;
                routeGrid.Rows.Clear();

                _lines.Clear();
                _relayRowIndex.Clear();
                _groupRowIndex.Clear();
                _hoverRowIndex = -1;

                if (lines.Count > 0)
                {
                    routeGrid.Rows.Add(lines.Count);
                    for (int i = 0; i < lines.Count; i++)
                    {
                        RouteLine line = lines[i];
                        _lines.Add(line);

                        if (line.Kind == RouteLineKind.Location) _groupRowIndex[line.Group] = i;
                        else _relayRowIndex[line.Row] = i;

                        ApplyLine(routeGrid.Rows[i], line);
                    }
                }
            }
            finally
            {
                routeGrid.ResumeLayout();
                _updatingGrid = false;
            }

            UpdateOverlay();
        }

        private void ApplyLine(DataGridViewRow gridRow, RouteLine line)
        {
            gridRow.Tag = line;
            gridRow.DefaultCellStyle.BackColor = BaseColorFor(line);

            // Painted by hand, but the values keep the cells meaningful to accessibility tools.
            gridRow.Cells[ColumnName].Value = line.Kind == RouteLineKind.Location
                ? line.Group.Pop.DisplayName
                : line.Row.Relay.Ipv4;
            gridRow.Cells[ColumnPing].Value = PingTextFor(line);

            var blocked = (DataGridViewCheckBoxCell)gridRow.Cells[ColumnBlocked];
            if (line.Kind == RouteLineKind.Location)
            {
                // Drawn by PaintGroupCheckbox so it can show a mixed state; clicks are handled
                // in CellMouseClick instead of by the cell itself.
                blocked.ThreeState = true;
                blocked.ReadOnly = true;
                blocked.Value = line.Group.AllBlocked
                    ? CheckState.Checked
                    : line.Group.AnyBlocked ? CheckState.Indeterminate : CheckState.Unchecked;
            }
            else
            {
                blocked.ThreeState = false;
                blocked.ReadOnly = false;
                blocked.Value = line.Row.IsBlocked;
            }
        }

        /// <summary>Refreshes one location row and any of its relay rows that are on screen.</summary>
        private void RefreshGroup(PopGroup group)
        {
            _updatingGrid = true;
            try
            {
                int index;
                if (_groupRowIndex.TryGetValue(group, out index))
                {
                    ApplyLine(routeGrid.Rows[index], _lines[index]);
                    routeGrid.InvalidateRow(index);
                }

                foreach (RouteRow row in group.Rows)
                {
                    if (!_relayRowIndex.TryGetValue(row, out index)) continue;
                    ApplyLine(routeGrid.Rows[index], _lines[index]);
                    routeGrid.InvalidateRow(index);
                }
            }
            finally
            {
                _updatingGrid = false;
            }
        }

        private void RefreshRelay(RouteRow row)
        {
            int index;
            if (_relayRowIndex.TryGetValue(row, out index)) routeGrid.InvalidateRow(index);
            if (_groupRowIndex.TryGetValue(row.Group, out index)) routeGrid.InvalidateRow(index);
        }

        private Color BaseColorFor(RouteLine line)
        {
            bool blocked = line.Kind == RouteLineKind.Location ? line.Group.AnyBlocked : line.Row.IsBlocked;
            return blocked ? BlockedTint : Color.Empty;
        }

        private void UpdateHeaders()
        {
            colName.HeaderText = "Route" + SortIndicator(RouteSort.Location);
            colPing.HeaderText = "Ping" + SortIndicator(RouteSort.Ping);
        }

        private string SortIndicator(RouteSort sort)
        {
            if (_sort != sort) return string.Empty;
            return _sortDescending ? "  ▼" : "  ▲";
        }

        private void UpdateOverlay()
        {
            if (!_routesLoaded)
            {
                ShowOverlay("Loading routes...");
                return;
            }

            if (_lines.Count == 0)
            {
                ShowOverlay(txtFilter.TextLength > 0
                    ? "No routes match \"" + txtFilter.Text.Trim() + "\"."
                    : "No routes to show.");
                return;
            }

            lblOverlay.Visible = false;
            routeGrid.Visible = true;
            btnClearFilter.Visible = txtFilter.TextLength > 0;
        }

        /// <summary>
        /// Swaps the grid out for a message. Hiding the grid rather than covering it keeps the
        /// two from depending on z-order.
        /// </summary>
        private void ShowOverlay(string text)
        {
            lblOverlay.Text = text;
            routeGrid.Visible = false;
            lblOverlay.Visible = true;
            btnClearFilter.Visible = txtFilter.TextLength > 0;
        }

        /// <summary>The DataGridView repaints the whole grid per cell change; buffering hides the flicker.</summary>
        private static void EnableDoubleBuffering(DataGridView grid)
        {
            PropertyInfo property = typeof(DataGridView).GetProperty(
                "DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            if (property != null) property.SetValue(grid, true, null);
        }

        #endregion

        #region Custom cell painting

        private void RouteGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                if (e.ColumnIndex == ColumnBlocked) PaintBlockedHeader(e);
                return;
            }

            RouteLine line = LineAt(e.RowIndex);
            if (line == null) return;

            if (e.ColumnIndex == ColumnName) PaintNameCell(e, line);
            else if (e.ColumnIndex == ColumnPing) PaintPingCell(e, line);
            else if (e.ColumnIndex == ColumnBlocked && line.Kind == RouteLineKind.Location) PaintGroupCheckbox(e, line.Group);
        }

        /// <summary>
        /// Locations get a chevron and a relay count; relays are indented under them and show
        /// the address that actually goes into the firewall rule.
        /// </summary>
        private void PaintNameCell(DataGridViewCellPaintingEventArgs e, RouteLine line)
        {
            e.PaintBackground(e.CellBounds, true);

            Graphics graphics = e.Graphics;
            Rectangle bounds = e.CellBounds;
            const TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix |
                                          TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding;

            if (line.Kind == RouteLineKind.Location)
            {
                if (line.Group.CanExpand)
                {
                    DrawChevron(graphics, bounds.Left + 14, bounds.Top + bounds.Height / 2, line.Group.IsExpanded);
                }

                var textBounds = Rectangle.FromLTRB(bounds.Left + 26, bounds.Top, bounds.Right - 6, bounds.Bottom);
                string name = line.Group.Pop.DisplayName;
                TextRenderer.DrawText(graphics, name, Font, textBounds, TextColor, flags);

                if (line.Group.CanExpand)
                {
                    Size nameSize = TextRenderer.MeasureText(graphics, name, Font, textBounds.Size, flags);
                    var countBounds = Rectangle.FromLTRB(
                        Math.Min(textBounds.Left + nameSize.Width + 8, bounds.Right - 6),
                        bounds.Top, bounds.Right - 6, bounds.Bottom);
                    string count = line.Group.Rows.Count.ToString(CultureInfo.CurrentCulture) + " relays";
                    TextRenderer.DrawText(graphics, count, Font, countBounds, MutedColor, flags);
                }
            }
            else
            {
                var textBounds = Rectangle.FromLTRB(bounds.Left + 44, bounds.Top, bounds.Right - 6, bounds.Bottom);
                TextRenderer.DrawText(graphics, line.Row.Relay.Ipv4, Font, textBounds, MutedColor, flags);
            }

            e.Handled = true;
        }

        private static void DrawChevron(Graphics graphics, int centreX, int centreY, bool expanded)
        {
            SmoothingMode previous = graphics.SmoothingMode;
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            try
            {
                Point[] points = expanded
                    ? new[]
                    {
                        new Point(centreX - 4, centreY - 2),
                        new Point(centreX + 4, centreY - 2),
                        new Point(centreX, centreY + 3)
                    }
                    : new[]
                    {
                        new Point(centreX - 2, centreY - 4),
                        new Point(centreX + 3, centreY),
                        new Point(centreX - 2, centreY + 4)
                    };

                using (var brush = new SolidBrush(MutedColor))
                {
                    graphics.FillPolygon(brush, points);
                }
            }
            finally
            {
                graphics.SmoothingMode = previous;
            }
        }

        /// <summary>Latency as a number plus a bar, so relative distance reads at a glance.</summary>
        private void PaintPingCell(DataGridViewCellPaintingEventArgs e, RouteLine line)
        {
            e.PaintBackground(e.CellBounds, true);

            Graphics graphics = e.Graphics;
            Rectangle bounds = e.CellBounds;

            long? milliseconds = PingValueFor(line);
            string text = PingTextFor(line);
            Color color = milliseconds.HasValue ? ColorForPing(milliseconds.Value) : DeadColor;

            var textBounds = Rectangle.FromLTRB(bounds.Left + 4, bounds.Top + 1, bounds.Right - 12, bounds.Bottom - 9);
            TextRenderer.DrawText(graphics, text, Font, textBounds, color,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            if (milliseconds.HasValue)
            {
                int left = bounds.Left + 10;
                int right = bounds.Right - 12;
                int width = right - left;
                if (width > 0)
                {
                    int top = bounds.Bottom - 8;
                    double fraction = Math.Min(1.0, Math.Max(0.02, milliseconds.Value / (double)PingBarScaleMs));
                    var filled = (int)Math.Round(width * fraction);

                    using (var track = new SolidBrush(TrackColor))
                    {
                        graphics.FillRectangle(track, left, top, width, 3);
                    }
                    using (var fill = new SolidBrush(color))
                    {
                        graphics.FillRectangle(fill, left, top, Math.Max(2, filled), 3);
                    }
                }
            }

            e.Handled = true;
        }

        /// <summary>A location's checkbox shows a mixed state when only some of its relays are blocked.</summary>
        private void PaintGroupCheckbox(DataGridViewCellPaintingEventArgs e, PopGroup group)
        {
            e.PaintBackground(e.CellBounds, true);

            CheckBoxState state = group.AllBlocked
                ? CheckBoxState.CheckedNormal
                : group.AnyBlocked ? CheckBoxState.MixedNormal : CheckBoxState.UncheckedNormal;

            DrawCentredCheckBox(e.Graphics, e.CellBounds, state);
            e.Handled = true;
        }

        private void PaintBlockedHeader(DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.CellBounds, false);

            CheckBoxState state = OverallBlockState();
            Size glyph = CheckBoxRenderer.GetGlyphSize(e.Graphics, CheckBoxState.UncheckedNormal);

            int left = e.CellBounds.Left + 10;
            int top = e.CellBounds.Top + (e.CellBounds.Height - glyph.Height) / 2;
            CheckBoxRenderer.DrawCheckBox(e.Graphics, new Point(left, top), state);

            var textBounds = Rectangle.FromLTRB(left + glyph.Width + 6, e.CellBounds.Top,
                e.CellBounds.Right, e.CellBounds.Bottom);
            TextRenderer.DrawText(e.Graphics, "Blocked", Font, textBounds,
                Color.FromArgb(90, 96, 106),
                TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);

            e.Handled = true;
        }

        private static void DrawCentredCheckBox(Graphics graphics, Rectangle bounds, CheckBoxState state)
        {
            Size glyph = CheckBoxRenderer.GetGlyphSize(graphics, CheckBoxState.UncheckedNormal);
            var origin = new Point(
                bounds.Left + (bounds.Width - glyph.Width) / 2,
                bounds.Top + (bounds.Height - glyph.Height) / 2);

            CheckBoxRenderer.DrawCheckBox(graphics, origin, state);
        }

        private CheckBoxState OverallBlockState()
        {
            bool any = false;
            bool all = _view.RelayCount > 0;

            foreach (RouteRow row in _view.Rows)
            {
                if (row.IsBlocked) any = true;
                else all = false;
            }

            return all ? CheckBoxState.CheckedNormal
                : any ? CheckBoxState.MixedNormal : CheckBoxState.UncheckedNormal;
        }

        private long? PingValueFor(RouteLine line)
        {
            if (line.Kind == RouteLineKind.Location) return line.Group.BestPing;

            PingResult? last = line.Row.LastPing;
            return last.HasValue && last.Value.Success ? (long?)last.Value.RoundtripMs : null;
        }

        private string PingTextFor(RouteLine line)
        {
            bool pinging = line.Kind == RouteLineKind.Location ? line.Group.IsPinging : line.Row.IsPinging;
            if (pinging) return "...";

            long? value = PingValueFor(line);
            if (value.HasValue)
            {
                return value.Value.ToString(CultureInfo.CurrentCulture) + " ms";
            }

            bool unreachable = line.Kind == RouteLineKind.Location
                ? line.Group.AllUnreachable
                : line.Row.LastPing.HasValue;

            return unreachable ? "no reply" : "—";
        }

        private Color ColorForPing(long milliseconds)
        {
            if (milliseconds <= _goodPingMs) return GoodColor;
            if (milliseconds <= _fairPingMs) return FairColor;
            return PoorColor;
        }

        #endregion

        #region Pinging

        private async Task PingVisibleRoutesAsync()
        {
            List<RouteRow> targets = VisibleRelays();
            if (targets.Count == 0) return;

            // Supersede any sweep still running, so repeated clicks do not pile up.
            CancellationTokenSource previous = _sweep;
            if (previous != null) previous.Cancel();

            var sweep = CancellationTokenSource.CreateLinkedTokenSource(_shutdown.Token);
            _sweep = sweep;

            btnPingRoutes.Enabled = false;
            BeginBusy(string.Format(CultureInfo.CurrentCulture, "Pinging {0} routes...", targets.Count));
            try
            {
                await PingRowsAsync(targets, sweep.Token);

                // Re-sorting only once the sweep finishes stops rows jumping under the cursor.
                if (_sort == RouteSort.Ping) RenderLines();
                SetStatus(DescribeSelection());
            }
            catch (OperationCanceledException)
            {
                // Superseded by a newer sweep, or the form is closing.
            }
            finally
            {
                EndBusy();
                if (!IsDisposed) btnPingRoutes.Enabled = true;
                if (ReferenceEquals(_sweep, sweep)) _sweep = null;
            }
        }

        /// <summary>
        /// Relays currently represented on screen: every relay of an expanded location, or just
        /// the first of a collapsed one, which is what its summary row reports.
        /// </summary>
        private List<RouteRow> VisibleRelays()
        {
            var targets = new List<RouteRow>();
            foreach (RouteLine line in _lines)
            {
                if (line.Kind == RouteLineKind.Relay)
                {
                    targets.Add(line.Row);
                }
                else if (!line.Group.IsExpanded && line.Group.Rows.Count > 0)
                {
                    targets.AddRange(line.Group.Rows);
                }
            }

            return targets;
        }

        /// <summary>
        /// Pings each relay with a bounded number of echoes in flight, updating every row as its
        /// reply arrives. The continuations run on the UI thread, so the grid writes are safe.
        /// </summary>
        private async Task PingRowsAsync(IList<RouteRow> rows, CancellationToken token)
        {
            using (var throttle = new SemaphoreSlim(_maxConcurrentPings))
            {
                var pings = new List<Task>(rows.Count);
                foreach (RouteRow row in rows)
                {
                    pings.Add(PingRowAsync(row, throttle, token));
                }

                await Task.WhenAll(pings);
            }
        }

        private async Task PingRowAsync(RouteRow row, SemaphoreSlim throttle, CancellationToken token)
        {
            await throttle.WaitAsync(token);
            try
            {
                if (IsDisposed) return;

                row.IsPinging = true;
                RefreshRelay(row);

                PingResult result = await PingService.SendAsync(row.Relay.Ipv4, _pingTimeoutMs, token);

                row.IsPinging = false;
                row.LastPing = result;
                if (!IsDisposed) RefreshRelay(row);
            }
            finally
            {
                row.IsPinging = false;
                throttle.Release();
            }
        }

        #endregion

        #region Blocking

        /// <summary>Writes the current selection for the given locations to the firewall.</summary>
        private async Task ApplyBlocksAsync(ICollection<PopGroup> groups)
        {
            if (groups.Count == 0) return;

            List<BlockRequest> requests = groups
                .Select(group => new BlockRequest(group.Pop.Code, group.BlockedRelays()))
                .ToList();

            BeginBusy("Updating firewall rules...");
            try
            {
                await Task.Run(() => _firewall.ApplyBlocks(requests), _shutdown.Token);
                SetStatus(DescribeSelection());
            }
            catch (OperationCanceledException)
            {
                // The form is closing.
            }
            catch (FirewallException ex)
            {
                SetStatus("The firewall rules could not be updated.");
                ShowError("The firewall rules could not be updated.", ex);
            }
            finally
            {
                EndBusy();
            }
        }

        private string DescribeSelection()
        {
            int relays = 0;
            int locations = 0;
            foreach (PopGroup group in _view.Groups)
            {
                int blocked = group.BlockedRelays().Count;
                if (blocked == 0) continue;

                relays += blocked;
                locations++;
            }

            if (relays == 0)
            {
                return string.Format(CultureInfo.CurrentCulture,
                    "{0} relays in {1} locations. Nothing blocked.", _view.RelayCount, _view.Groups.Count);
            }

            return string.Format(CultureInfo.CurrentCulture,
                "Blocking {0} of {1} relays across {2} location(s).", relays, _view.RelayCount, locations);
        }

        private async Task SetGroupBlockedAsync(PopGroup group, bool blocked)
        {
            group.SetAllBlocked(blocked);
            RefreshGroup(group);
            routeGrid.InvalidateCell(ColumnBlocked, -1);
            await ApplyBlocksAsync(new[] { group });
        }

        #endregion

        #region Grid events

        private RouteLine LineAt(int rowIndex)
        {
            return rowIndex >= 0 && rowIndex < _lines.Count ? _lines[rowIndex] : null;
        }

        /// <summary>Commits a checkbox the moment it is clicked, so CellValueChanged sees the new value.</summary>
        private void RouteGrid_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (routeGrid.IsCurrentCellDirty &&
                routeGrid.CurrentCell != null &&
                routeGrid.CurrentCell.ColumnIndex == ColumnBlocked)
            {
                routeGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private async void RouteGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_updatingGrid || e.ColumnIndex != ColumnBlocked || e.RowIndex < 0) return;

            RouteLine line = LineAt(e.RowIndex);
            if (line == null || line.Kind != RouteLineKind.Relay) return;

            var value = routeGrid.Rows[e.RowIndex].Cells[ColumnBlocked].Value as bool?;
            line.Row.IsBlocked = value.GetValueOrDefault();

            RefreshGroup(line.Group);
            routeGrid.InvalidateCell(ColumnBlocked, -1);
            await ApplyBlocksAsync(new[] { line.Group });
        }

        /// <summary>Handles the location checkbox, which is painted rather than edited.</summary>
        private async void RouteGrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != ColumnBlocked || e.Button != MouseButtons.Left) return;

            RouteLine line = LineAt(e.RowIndex);
            if (line == null || line.Kind != RouteLineKind.Location) return;

            await SetGroupBlockedAsync(line.Group, !line.Group.AnyBlocked);
        }

        private async void RouteGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            RouteLine line = LineAt(e.RowIndex);
            if (line == null) return;

            try
            {
                if (e.ColumnIndex == ColumnName && line.Kind == RouteLineKind.Location)
                {
                    ToggleExpansion(line.Group);
                }
                else if (e.ColumnIndex == ColumnPing)
                {
                    await PingLineAsync(line);
                }
            }
            catch (OperationCanceledException)
            {
                // The form is closing.
            }
        }

        private async Task PingLineAsync(RouteLine line)
        {
            List<RouteRow> targets = line.Kind == RouteLineKind.Location
                ? new List<RouteRow>(line.Group.Rows)
                : new List<RouteRow> { line.Row };

            await PingRowsAsync(targets, _shutdown.Token);
        }

        /// <summary>Clicking a header sorts by it; the Blocked header blocks or unblocks everything.</summary>
        private async void RouteGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == ColumnBlocked)
            {
                if (_view.RelayCount == 0) return;

                bool block = OverallBlockState() != CheckBoxState.CheckedNormal;
                foreach (PopGroup group in _view.Groups) group.SetAllBlocked(block);

                RenderLines();
                routeGrid.InvalidateCell(ColumnBlocked, -1);
                await ApplyBlocksAsync(_view.Groups);
                return;
            }

            RouteSort requested = e.ColumnIndex == ColumnPing ? RouteSort.Ping : RouteSort.Location;
            if (_sort == requested) _sortDescending = !_sortDescending;
            else
            {
                _sort = requested;
                _sortDescending = false;
            }

            UpdateHeaders();
            RenderLines();
        }

        private void ToggleExpansion(PopGroup group)
        {
            if (!group.CanExpand) return;

            group.IsExpanded = !group.IsExpanded;
            RenderLines();

            // Expanding reveals relays that may never have been pinged.
            if (group.IsExpanded)
            {
                var pending = group.Rows.Where(row => !row.LastPing.HasValue).ToList();
                if (pending.Count > 0)
                {
                    Task ignored = PingRowsAsync(pending, _shutdown.Token);
                    GC.KeepAlive(ignored);
                }
            }
        }

        private void RouteGrid_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            SetHoveredRow(e.RowIndex);
        }

        private void RouteGrid_MouseLeave(object sender, EventArgs e)
        {
            SetHoveredRow(-1);
        }

        private void SetHoveredRow(int rowIndex)
        {
            if (rowIndex == _hoverRowIndex) return;

            if (_hoverRowIndex >= 0 && _hoverRowIndex < routeGrid.RowCount)
            {
                RouteLine previous = LineAt(_hoverRowIndex);
                if (previous != null) routeGrid.Rows[_hoverRowIndex].DefaultCellStyle.BackColor = BaseColorFor(previous);
            }

            _hoverRowIndex = rowIndex >= 0 && rowIndex < routeGrid.RowCount ? rowIndex : -1;

            if (_hoverRowIndex >= 0 && LineAt(_hoverRowIndex) != null)
            {
                routeGrid.Rows[_hoverRowIndex].DefaultCellStyle.BackColor = HoverColor;
            }
        }

        private async void RouteGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (routeGrid.CurrentRow == null) return;

            RouteLine line = LineAt(routeGrid.CurrentRow.Index);
            if (line == null) return;

            if (e.KeyCode == Keys.Enter && line.Kind == RouteLineKind.Location)
            {
                ToggleExpansion(line.Group);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Right && line.Kind == RouteLineKind.Location && !line.Group.IsExpanded)
            {
                ToggleExpansion(line.Group);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Left && line.Kind == RouteLineKind.Location && line.Group.IsExpanded)
            {
                ToggleExpansion(line.Group);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Space && routeGrid.CurrentCell != null &&
                     routeGrid.CurrentCell.ColumnIndex != ColumnBlocked)
            {
                e.Handled = true;
                await ToggleBlockAsync(line);
            }
        }

        private async Task ToggleBlockAsync(RouteLine line)
        {
            if (line.Kind == RouteLineKind.Location)
            {
                await SetGroupBlockedAsync(line.Group, !line.Group.AnyBlocked);
                return;
            }

            line.Row.IsBlocked = !line.Row.IsBlocked;
            RefreshGroup(line.Group);
            routeGrid.InvalidateCell(ColumnBlocked, -1);
            await ApplyBlocksAsync(new[] { line.Group });
        }

        #endregion

        #region Context menu

        private void RowMenu_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Point local = routeGrid.PointToClient(Cursor.Position);
            DataGridView.HitTestInfo hit = routeGrid.HitTest(local.X, local.Y);

            _menuLine = LineAt(hit.RowIndex);
            if (_menuLine == null)
            {
                e.Cancel = true;
                return;
            }

            routeGrid.CurrentCell = routeGrid.Rows[hit.RowIndex].Cells[ColumnName];

            bool isLocation = _menuLine.Kind == RouteLineKind.Location;
            menuCopyAddress.Text = isLocation && _menuLine.Group.Rows.Count > 1
                ? "Copy IP addresses"
                : "Copy IP address";
            menuPingRow.Text = isLocation ? "Ping this location" : "Ping this relay";
            menuToggleBlock.Text = (isLocation ? _menuLine.Group.AnyBlocked : _menuLine.Row.IsBlocked)
                ? "Unblock"
                : "Block";
        }

        private void MenuCopyAddress_Click(object sender, EventArgs e)
        {
            if (_menuLine == null) return;

            var text = new StringBuilder();
            if (_menuLine.Kind == RouteLineKind.Location)
            {
                foreach (RouteRow row in _menuLine.Group.Rows) text.AppendLine(row.Relay.Ipv4);
            }
            else
            {
                text.Append(_menuLine.Row.Relay.Ipv4);
            }

            try
            {
                Clipboard.SetText(text.ToString().TrimEnd());
                SetStatus("Copied to the clipboard.");
            }
            catch (ExternalException ex)
            {
                Debug.WriteLine("Clipboard unavailable: " + ex.Message);
                SetStatus("The clipboard was unavailable.");
            }
        }

        private async void MenuPingRow_Click(object sender, EventArgs e)
        {
            if (_menuLine == null) return;

            try
            {
                await PingLineAsync(_menuLine);
            }
            catch (OperationCanceledException)
            {
                // The form is closing.
            }
        }

        private async void MenuToggleBlock_Click(object sender, EventArgs e)
        {
            if (_menuLine == null) return;
            await ToggleBlockAsync(_menuLine);
        }

        #endregion

        #region Filter and buttons

        private void TxtFilter_TextChanged(object sender, EventArgs e)
        {
            if (!_routesLoaded) return;
            RenderLines();
        }

        private void BtnClearFilter_Click(object sender, EventArgs e)
        {
            txtFilter.Clear();
            txtFilter.Focus();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.F))
            {
                txtFilter.Focus();
                txtFilter.SelectAll();
                return true;
            }

            if (keyData == Keys.F5 && btnPingRoutes.Enabled)
            {
                Task ignored = PingVisibleRoutesAsync();
                GC.KeepAlive(ignored);
                return true;
            }

            if (keyData == Keys.Escape && txtFilter.Focused && txtFilter.TextLength > 0)
            {
                txtFilter.Clear();
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private async void BtnPingRoutes_Click(object sender, EventArgs e)
        {
            await PingVisibleRoutesAsync();
        }

        /// <summary>Re-opens the app id prompt and reloads the routes for the chosen game.</summary>
        private async void BtnChangeGame_Click(object sender, EventArgs e)
        {
            int chosen;
            using (var prompt = new AppIdPromptForm(_appId))
            {
                if (prompt.ShowDialog(this) != DialogResult.OK) return;
                chosen = prompt.AppId;
            }

            if (chosen == _appId) return;

            btnChangeGame.Enabled = false;
            try
            {
                await LoadGameAsync(chosen);
            }
            catch (OperationCanceledException)
            {
                // The form is closing.
            }
            catch (SdrConfigException ex)
            {
                UpdateOverlay();
                SetStatus(string.Format(CultureInfo.CurrentCulture,
                    "App id {0} could not be loaded; still showing app id {1}.", chosen, _appId));
                ShowError("Could not load that app id.", ex);
            }
            finally
            {
                if (!IsDisposed) btnChangeGame.Enabled = true;
            }
        }

        private async void BtnClearRules_Click(object sender, EventArgs e)
        {
            btnClearRules.Enabled = false;
            BeginBusy("Clearing firewall rules...");
            try
            {
                int removed = await Task.Run(
                    () => _firewall.RemoveRulesWithPrefix(FirewallService.RulePrefix), _shutdown.Token);

                foreach (PopGroup group in _view.Groups) group.SetAllBlocked(false);
                RenderLines();

                SetStatus(string.Format(CultureInfo.CurrentCulture,
                    "Removed {0} firewall rule(s) created by this tool.", removed));
            }
            catch (OperationCanceledException)
            {
                // The form is closing.
            }
            catch (FirewallException ex)
            {
                SetStatus("The firewall rules could not be cleared.");
                ShowError("The firewall rules could not be cleared.", ex);
            }
            finally
            {
                EndBusy();
                if (!IsDisposed) btnClearRules.Enabled = true;
            }
        }

        private void BtnAbout_Click(object sender, EventArgs e)
        {
            string name = KnownGames.NameFor(_appId);
            string message = string.Join(Environment.NewLine,
                "SteamRouteTool " + ProductVersion,
                string.Empty,
                "Currently showing: " + (name ?? "app id " + _appId.ToString(CultureInfo.CurrentCulture)),
                "Steam app ID: " + _appId.ToString(CultureInfo.CurrentCulture),
                "Rules are named \"" + FirewallService.RulePrefix + "<protocol>-<location>\".",
                string.Empty,
                "Shortcuts: Ctrl+F filter, F5 re-ping, Space block, Enter expand.",
                string.Empty,
                "Created by Froody.");

            MessageBox.Show(this, message, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region Status and shutdown

        private void BeginBusy(string status)
        {
            _busyDepth++;
            progressStatus.Visible = true;
            SetStatus(status);
        }

        private void EndBusy()
        {
            if (_busyDepth > 0) _busyDepth--;
            if (_busyDepth == 0 && !IsDisposed) progressStatus.Visible = false;
        }

        private void SetStatus(string status)
        {
            if (IsDisposed) return;
            lblStatus.Text = status ?? string.Empty;
        }

        private void ShowError(string caption, Exception error)
        {
            if (IsDisposed) return;

            string detail = error.InnerException != null
                ? error.Message + Environment.NewLine + Environment.NewLine + error.InnerException.Message
                : error.Message;

            MessageBox.Show(this, detail, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Stops in-flight pings and firewall calls from resuming into a disposed form.
            _shutdown.Cancel();
            base.OnFormClosing(e);
        }

        #endregion
    }
}
