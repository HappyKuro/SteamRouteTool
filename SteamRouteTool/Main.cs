using NetFwTypeLib;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SteamRouteTool
{
    public partial class Main : Form
    {
        public List<Route> routes = new List<Route>();
        int rowCount = 0;
        bool columnChecked = false;
        bool firstLoad = true;
        string networkconfigURL = @"https://api.steampowered.com/ISteamApps/GetSDRConfig/v1?appid=440";

        public Main()
        {
            InitializeComponent();
            ClearTF2RoutingToolRules();

            Task.Run(() => PopulateRoutesAsync());
        }

        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || this.Disposing) return;
            try
            {
                if (this.InvokeRequired) this.Invoke(action);
                else action();
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void ClearTF2RoutingToolRules()
        {
            try
            {
                Type tNetFwPolicy2 = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                INetFwPolicy2 fwPolicy2 = (INetFwPolicy2)Activator.CreateInstance(tNetFwPolicy2);
                var toRemove = new List<string>();
                foreach (INetFwRule rule in fwPolicy2.Rules)
                {
                    if (rule.Name != null && rule.Name.Contains("TF2RoutingTool-")) toRemove.Add(rule.Name);
                }
                foreach (var name in toRemove)
                {
                    try { fwPolicy2.Rules.Remove(name); }
                    catch (Exception ex) { Debug.WriteLine($"ClearTF2RoutingToolRules: failed removing {name}: {ex}"); }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ClearTF2RoutingToolRules: " + ex);
            }

            SafeInvoke(() =>
            {
                MessageBox.Show(
                    "Welcome to SteamRouteTool!" + Environment.NewLine +
                    "Last changed: 15/11/2024" + Environment.NewLine +
                    "Improved firewallpolicy, fwPolicy2." + Environment.NewLine +
                    "Added MessageBox welcome." + Environment.NewLine +
                    "changed networkconfigURL from 730 to 440.",
                    "Welcome!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            });
        }

        private async Task PopulateRoutesAsync()
        {
            string raw = null;
            try
            {
                using (var wc = new WebClient())
                {
                    raw = await wc.DownloadStringTaskAsync(networkconfigURL);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PopulateRoutesAsync: download failed: " + ex);
                SafeInvoke(() => lb_GettingRoutes.Visible = false);
                return;
            }

            JObject jObj;
            try
            {
                jObj = JsonConvert.DeserializeObject<JObject>(raw);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PopulateRoutesAsync: json parse failed: " + ex);
                SafeInvoke(() => lb_GettingRoutes.Visible = false);
                return;
            }

            var pops = jObj["pops"] as JObject;
            if (pops == null)
            {
                SafeInvoke(() => lb_GettingRoutes.Visible = false);
                return;
            }

            foreach (var rc in pops)
            {
                var value = rc.Value as JObject;
                if (value == null) continue;

                if (!value.ContainsKey("relays")) continue;
                if (value.ToString().Contains("cloud-test")) continue;

                var route = new Route
                {
                    name = rc.Key,
                    ranges = new Dictionary<string, string>(),
                    row_index = new List<int>(),
                    extended = false,
                    all_check = false,
                    pw = value.ToString().Contains("partners\": 2")
                };

                if (value.ContainsKey("desc"))
                {
                    route.desc = value["desc"]?.ToString();
                }

                foreach (var rangeToken in value["relays"].Children<JObject>())
                {
                    var ipv4 = rangeToken["ipv4"]?.ToString();
                    var portRange = rangeToken["port_range"] != null ? rangeToken["port_range"].ToString() : string.Empty;
                    if (string.IsNullOrWhiteSpace(ipv4)) continue;

                    Debug.WriteLine($"{ipv4} {portRange}");
                    route.ranges.Add(ipv4, portRange);
                    route.row_index.Add(rowCount);
                    rowCount++;
                }

                routes.Add(route);
            }

            SafeInvoke(() =>
            {
                btn_PingRoutes.Enabled = true;
                lb_GettingRoutes.Visible = false;
                btn_About.Visible = true;
            });

            SafeInvoke(() =>
            {
                PopulateRouteDataGrid();
            });
        }

        private void PopulateRouteDataGrid()
        {
            if (routeDataGrid.RowCount == 0)
            {
                for (int i = 0; i < rowCount; i++)
                {
                    routeDataGrid.Rows.Add();
                    routeDataGrid.Rows[i].Cells[2].Value = false;
                }
            }

            foreach (Route route in routes)
            {
                for (int i = 0; i < route.ranges.Count; i++)
                {
                    var label = route.desc != null ? $"{route.desc} {i + 1}" : $"{route.name} {i + 1}";
                    routeDataGrid.Rows[route.row_index[i]].Cells[0].Value = label;

                    if (!route.extended)
                    {
                        var firstLabel = route.desc != null ? route.desc : route.name;
                        routeDataGrid.Rows[route.row_index[0]].Cells[0].Value = firstLabel;
                    }

                    routeDataGrid.Rows[route.row_index[i]].Visible = (i == 0) || route.extended;
                }
            }

            if (firstLoad)
            {
                Task.Run(() => PingRoutesAsync());
                Task.Run(() => GetCurrentBlockedAsync());
            }
            firstLoad = false;
        }

        private async Task PingSingleRouteAsync(Route route)
        {
            await Task.Run(() =>
            {
                for (int i = 0; i < route.ranges.Count; i++)
                {
                    var rowIdx = route.row_index[i];
                    SafeInvoke(() => routeDataGrid.Rows[rowIdx].Cells[1].Style.BackColor = Color.Black);

                    string responseTime = PingHost(route.ranges.Keys.ElementAt(i));
                    SafeInvoke(() =>
                    {
                        if (responseTime != "UNK")
                        {
                            int ms;
                            int.TryParse(responseTime, out ms);
                            if (ms <= 50) routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.Green;
                            else if (ms <= 100) routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.Orange;
                            else routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.Red;
                        }
                        else
                        {
                            routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.DarkRed;
                        }

                        routeDataGrid.Rows[rowIdx].Cells[1].Value = responseTime;
                        routeDataGrid.Rows[rowIdx].Cells[1].Style.BackColor = Color.White;
                    });
                }
            });
        }

        private async Task PingRoutesAsync()
        {
            var tasks = new List<Task>();
            foreach (Route route in routes)
            {
                tasks.Add(Task.Run(() =>
                {
                    for (int i = 0; i < route.ranges.Count; i++)
                    {
                        var rowIdx = route.row_index[i];
                        SafeInvoke(() => routeDataGrid.Rows[rowIdx].Cells[1].Style.BackColor = Color.Black);

                        string responseTime = PingHost(route.ranges.Keys.ElementAt(i));

                        SafeInvoke(() =>
                        {
                            if (responseTime != "BLK")
                            {
                                int ms;
                                int.TryParse(responseTime, out ms);
                                if (ms <= 100) routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.Green;
                                else if (ms <= 100) routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.Orange;
                                else routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.Red;
                            }
                            else
                            {
                                routeDataGrid.Rows[rowIdx].Cells[1].Style.ForeColor = Color.DarkRed;
                            }

                            routeDataGrid.Rows[rowIdx].Cells[1].Value = responseTime;
                            routeDataGrid.Rows[rowIdx].Cells[1].Style.BackColor = Color.White;
                        });
                    }
                }));
            }

            await Task.WhenAll(tasks);
        }

        public static string PingHost(string host)
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = ping.Send(host, 1000);
                    if (reply != null && reply.Status == IPStatus.Success && reply.RoundtripTime > 0)
                    {
                        return reply.RoundtripTime.ToString();
                    }

                    return "BLK";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PingHost: " + ex);
                return "BLK";
            }
        }

        private Task GetCurrentBlockedAsync()
        {
            return Task.Run(() =>
            {
                try
                {
                    Type tNetFwPolicy2 = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                    INetFwPolicy2 fwPolicy2 = (INetFwPolicy2)Activator.CreateInstance(tNetFwPolicy2);

                    foreach (INetFwRule rule in fwPolicy2.Rules)
                    {
                        if (rule.Name == null) continue;
                        if (!rule.Name.Contains("SteamRouteTool-")) continue;

                        var parts = rule.Name.Split('-');
                        if (parts.Length < 3) continue;
                        string name = parts[2];

                        var addr = new List<string>();
                        if (string.IsNullOrWhiteSpace(rule.RemoteAddresses)) continue;
                        foreach (string tosplit in rule.RemoteAddresses.Split(','))
                        {
                            var a = tosplit.Split('/')[0];
                            if (!string.IsNullOrWhiteSpace(a)) addr.Add(a);
                        }

                        foreach (Route route in routes)
                        {
                            if (route.name != name) continue;

                            bool extended = true;
                            bool firstBlocked = false;
                            int blockedCount = 0;
                            for (int i = 0; i < route.ranges.Count; i++)
                            {
                                if (addr.Contains(route.ranges.Keys.ElementAt(i)))
                                {
                                    int rowIdx = route.row_index[i];
                                    SafeInvoke(() => routeDataGrid.Rows[rowIdx].Cells[2].Value = true);
                                    if (i != 0) blockedCount++;
                                    if (i == 0) firstBlocked = true;
                                }
                            }

                            if (blockedCount == route.ranges.Count - 1 && firstBlocked) extended = false;
                            route.extended = extended;
                            if (extended)
                            {
                                foreach (int index in route.row_index) SafeInvoke(() => routeDataGrid.Rows[index].Visible = true);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("GetCurrentBlockedAsync: " + ex);
                }
            });
        }

        private void Btn_ClearRules_Click(object sender, EventArgs e)
        {
            try
            {
                Type tNetFwPolicy2 = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                INetFwPolicy2 fwPolicy2 = (INetFwPolicy2)Activator.CreateInstance(tNetFwPolicy2);

                var toRemove = new List<string>();
                foreach (INetFwRule rule in fwPolicy2.Rules)
                {
                    if (rule.Name != null && rule.Name.Contains("SteamRouteTool-")) toRemove.Add(rule.Name);
                }

                foreach (var name in toRemove)
                {
                    try { fwPolicy2.Rules.Remove(name); }
                    catch (Exception ex) { Debug.WriteLine("Btn_ClearRules_Click remove: " + ex); }
                }

                SafeInvoke(() =>
                {
                    for (int i = 0; i < routeDataGrid.Rows.Count; i++) routeDataGrid.Rows[i].Cells[2].Value = false;
                    MessageBox.Show("You have cleared all firewall rules created by this tool.", "Steam Route Tool - Rules Clear");
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Btn_ClearRules_Click: " + ex);
                SafeInvoke(() => MessageBox.Show("Failed to clear rules. Check permissions.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error));
            }
        }

        private void Btn_PingRoutes_Click(object sender, EventArgs e)
        {
            Task.Run(() => PingRoutesAsync());
        }

        void RouteDataGrid_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (routeDataGrid.CurrentCell != null && routeDataGrid.CurrentCell.ColumnIndex == 2 && routeDataGrid.IsCurrentCellDirty)
            {
                routeDataGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void RouteDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < -1) return;

            if (e.ColumnIndex == 0 && e.RowIndex != -1)
            {
                var currentRoute = routes.First(x => x.row_index.Contains(e.RowIndex));

                if (currentRoute.all_check && !currentRoute.extended)
                {
                    bool blocked = false;
                    for (int i = 0; i < currentRoute.row_index.Count; i++)
                    {
                        if (i == 0)
                        {
                            blocked = Convert.ToBoolean(routeDataGrid.Rows[currentRoute.row_index[i]].Cells[2].Value);
                        }
                        else
                        {
                            routeDataGrid.Rows[currentRoute.row_index[i]].Cells[2].Value = blocked;
                        }
                    }
                }

                currentRoute.extended = !currentRoute.extended;
                currentRoute.all_check = !currentRoute.all_check;

                PopulateRouteDataGrid();
                Task.Run(() => PingSingleRouteAsync(currentRoute));
            }

            if (e.ColumnIndex == 1 && e.RowIndex != -1)
            {
                Task.Run(() =>
                {
                    var currentRoute = routes.First(x => x.row_index.Contains(e.RowIndex));
                    for (int i = 0; i < currentRoute.row_index.Count; i++)
                    {
                        if (currentRoute.row_index[i] == e.RowIndex)
                        {
                            SafeInvoke(() => routeDataGrid.Rows[e.RowIndex].Cells[1].Style.BackColor = Color.Black);
                            string responseTime = PingHost(currentRoute.ranges.Keys.ElementAt(i));
                            SafeInvoke(() =>
                            {
                                if (responseTime != "UNK")
                                {
                                    int ms;
                                    int.TryParse(responseTime, out ms);
                                    if (ms <= 75) routeDataGrid.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;
                                    else if (ms <= 100) routeDataGrid.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Orange;
                                    else routeDataGrid.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Red;
                                }
                                else
                                {
                                    routeDataGrid.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.DarkRed;
                                }

                                routeDataGrid.Rows[e.RowIndex].Cells[1].Value = responseTime;
                                routeDataGrid.Rows[e.RowIndex].Cells[1].Style.BackColor = Color.White;
                            });
                        }
                    }
                });
            }

            if (e.ColumnIndex == 2 && e.RowIndex != -1)
            {
                var currentRoute = routes.First(x => x.row_index.Contains(e.RowIndex));
                currentRoute.all_check = !currentRoute.extended;
                Task.Run(() => SetRule(currentRoute));
            }

            if (e.ColumnIndex == 2 && e.RowIndex == -1)
            {
                columnChecked = !columnChecked;
                for (int i = 0; i < routeDataGrid.Rows.Count; i++)
                {
                    routeDataGrid.Rows[i].Cells[2].Value = columnChecked;
                }

                if (columnChecked)
                {
                    foreach (var route in routes) Task.Run(() => SetRule(route));
                }
                else
                {
                    Task.Run(() =>
                    {
                        try
                        {
                            Type tNetFwPolicy2 = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                            INetFwPolicy2 fwPolicy2 = (INetFwPolicy2)Activator.CreateInstance(tNetFwPolicy2);
                            foreach (var route in routes)
                            {
                                try
                                {
                                    fwPolicy2.Rules.Remove("SteamRouteTool-TCP-" + route.name);
                                    fwPolicy2.Rules.Remove("SteamRouteTool-UDP-" + route.name);
                                    fwPolicy2.Rules.Remove("SteamRouteTool-ICMP-" + route.name);
                                }
                                catch (Exception ex) { Debug.WriteLine("Remove rules for route: " + ex); }
                            }
                        }
                        catch (Exception ex) { Debug.WriteLine("Clear all rules: " + ex); }
                    });
                }
            }
        }

        private void SetRule(Route route)
        {
            try
            {
                Type tNetFwPolicy2 = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                INetFwPolicy2 fwPolicy2 = (INetFwPolicy2)Activator.CreateInstance(tNetFwPolicy2);
                try
                {
                    fwPolicy2.Rules.Remove("SteamRouteTool-TCP-" + route.name);
                    fwPolicy2.Rules.Remove("SteamRouteTool-UDP-" + route.name);
                    fwPolicy2.Rules.Remove("SteamRouteTool-ICMP-" + route.name);
                }
                catch (Exception ex) { Debug.WriteLine("SetRule: remove existing rules: " + ex); }

                string remoteAddresses = "";

                for (int i = 0; i < route.ranges.Count; i++)
                {
                    var isChecked = false;
                    try
                    {
                        var cellVal = routeDataGrid.Rows[route.row_index[i]].Cells[2].Value;
                        isChecked = cellVal != null && Convert.ToBoolean(cellVal);
                    }
                    catch { isChecked = false; }

                    if (!isChecked) continue;

                    if (i == 0 && route.all_check)
                    {
                        foreach (var kv in route.ranges)
                        {
                            remoteAddresses += kv.Key + ",";
                        }
                        break;
                    }
                    else
                    {
                        remoteAddresses += route.ranges.Keys.ElementAt(i) + ",";
                    }
                }

                if (!string.IsNullOrEmpty(remoteAddresses))
                {
                    remoteAddresses = remoteAddresses.TrimEnd(',');

                    // UDP Rule
                    INetFwRule2 udpRule = (INetFwRule2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                    udpRule.Enabled = true;
                    udpRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
                    udpRule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
                    udpRule.RemoteAddresses = remoteAddresses;
                    udpRule.Protocol = 17;
                    udpRule.RemotePorts = "27015-27202";
                    udpRule.Name = "SteamRouteTool-UDP-" + route.name;

                    // TCP Rule
                    INetFwRule2 tcpRule = (INetFwRule2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                    tcpRule.Enabled = true;
                    tcpRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
                    tcpRule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
                    tcpRule.RemoteAddresses = remoteAddresses;
                    tcpRule.Protocol = 6;
                    tcpRule.RemotePorts = "27015-27202";
                    tcpRule.Name = "SteamRouteTool-TCP-" + route.name;

                    // ICMP Rule
                    INetFwRule2 icmpRule = (INetFwRule2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                    icmpRule.Enabled = true;
                    icmpRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
                    icmpRule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
                    icmpRule.RemoteAddresses = remoteAddresses;
                    icmpRule.Protocol = 1;
                    icmpRule.Name = "SteamRouteTool-ICMP-" + route.name;

                    INetFwPolicy2 firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
                    firewallPolicy.Rules.Add(udpRule);
                    firewallPolicy.Rules.Add(tcpRule);
                    firewallPolicy.Rules.Add(icmpRule);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("SetRule: " + ex);
            }
        }

        private void Btn_About_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Version: " + ProductVersion + Environment.NewLine + "SteamRouteTool is created by Froody.", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}