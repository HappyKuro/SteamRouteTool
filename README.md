# SteamRouteTool

Blocks individual Steam Datagram Relay (SDR) routes so games that matchmake through Valve's
relay network stop connecting you to data centres you don't want.

Steam publishes, per app, the list of relay clusters ("points of presence") and their IP
addresses. SteamRouteTool reads that list, pings every relay, and writes outbound Windows
Firewall rules for the ones you tick. Valve's matchmaking then routes you elsewhere.

> Blocking routes changes which servers you can be matched to and can increase queue times.
> Everything the tool writes is reversible from inside the app with **Clear Rules**.

## Requirements

* Windows with the Windows Defender Firewall service running
* .NET Framework 4.7.2
* **Administrator rights** — adding or removing firewall rules requires elevation. The app
  requests this automatically; accept the UAC prompt.

## Getting started

1. Launch `SteamRouteTool.exe` and accept the UAC prompt.
2. **Choose a game.** The tool asks which Steam app to work on before the main window opens:

   ```
   Which game do you want to work on?
   [ Team Fortress 2 (440)                     v ]
   ```

   Pick one from the list, or type any Steam app ID. Your choice is remembered for next time.
3. The main window lists every relay location for that app, with its latency.

### Which app ID should I use?

The list covers the games this is most often used with:

| Game | App ID |
| --- | --- |
| Team Fortress 2 | 440 |
| Counter-Strike 2 | 730 |
| Dota 2 | 570 |
| Deadlock | 1422450 |
| Rust | 252490 |

The list is a convenience, not a restriction — **any Steam app ID works**. If you're not sure
of one, it's the number in the game's store URL:
`store.steampowered.com/app/`**`440`**`/Team_Fortress_2/`.

Note that Steam serves a shared default relay list for most apps; a few (Counter-Strike 2,
Dota 2 and Rust among them) publish their own larger set. If an app ID isn't valid, the tool
says so and keeps showing the game you already had loaded.

## Using it

Each location is one row, showing its relay count and its best latency. Click the row (or the
chevron) to expand it and see the individual relay addresses underneath.

| Column | Click it to |
| --- | --- |
| **Route** | Expand or collapse the location |
| **Ping** | Re-ping that location or relay |
| **Blocked** | Block or unblock it |

Latency shows as a number plus a bar, so you can compare distances at a glance, coloured
green / orange / red against the thresholds in the config file. `no reply` means the relay
didn't answer; `—` means it hasn't been pinged yet.

Ticking a **location** blocks every relay there. Expand it first to block relays individually
— the location's checkbox then shows a **mixed** mark, meaning "some but not all". The
checkbox in the **Blocked** column header does the same for every location at once.

**Sort** by clicking the *Route* or *Ping* header; click again to reverse. Sorting by ping is
the quick way to find your nearest data centres. Relays always stay grouped under their
location, and the order only settles once a ping sweep finishes, so rows don't jump around
under the cursor.

**Filter** with the box at the top. It matches the location name, its short Valve code, and
relay IP addresses — so typing `162.254` shows every location using that block.

**Right-click** any row to copy its IP address (or all of a location's addresses), ping it, or
block it.

### Buttons and shortcuts

| Button | What it does |
| --- | --- |
| **Ping Routes** | Re-pings everything on screen |
| **Clear Rules** | Removes every firewall rule this tool created |
| **Change Game** | Re-opens the app ID prompt and reloads without restarting |
| **About** | Version and the app ID currently loaded |

| Key | Action |
| --- | --- |
| `Ctrl+F` | Jump to the filter box |
| `Esc` | Clear the filter |
| `F5` | Re-ping |
| `Enter` / `←` / `→` | Expand or collapse the selected location |
| `Space` | Block or unblock the selected row |

## How blocking works

For each location you block, three outbound **block** rules are added:

```
SteamRouteTool-UDP-<location>
SteamRouteTool-TCP-<location>
SteamRouteTool-ICMP-<location>
```

Each targets the relay addresses for that location, on the port range Steam publishes for
those relays (it differs per location — 27015–27060, 27015–27078 and 27015–27140 are all in
current use). You can inspect them in `wf.msc` under the **SteamRouteTool** group.

Two things worth knowing:

* Blocked routes stop replying to ping, so their latency shows as `-`. That's the ICMP rule
  working, not a fault.
* Rules apply machine-wide to those addresses, not only to the game.

On start-up the tool reads existing rules back, so the checkboxes always reflect what is
really in the firewall — including after a restart, where a partly-blocked location comes back
with its mixed mark. It also clears any rules left behind by the older **TF2RoutingTool**, so
you don't have to go back and do it there.

## Configuration

Optional settings in `SteamRouteTool.exe.config`:

| Key | Default | Meaning |
| --- | --- | --- |
| `appId` | `440` | App ID offered on a first run, before the tool remembers your choice |
| `pingTimeoutMs` | `1000` | How long to wait for an ICMP reply |
| `maxConcurrentPings` | `16` | How many relays are pinged at once |
| `goodPingMs` | `50` | At or below this, latency is green |
| `fairPingMs` | `100` | At or below this, orange; above it, red |

## Building

Open `SteamRouteTool.sln` in Visual Studio and build, or from a developer prompt:

```
msbuild SteamRouteTool.sln -t:Rebuild -p:Configuration=Release
```

Targets .NET Framework 4.7.2. `Newtonsoft.Json` restores from `packages.config`; the firewall
API is the `NetFwTypeLib` COM interop that ships with Windows.

Layout:

```
Models/      Relay, PointOfPresence, PortRange, PingResult, SteamGame  - plain data
Services/    SdrConfigClient (download + parse)
             PingService     (async ICMP)
             FirewallService (all COM firewall access)
ViewModel/   RouteView       - grouping, filtering, sorting, block/expand state
MainForm     UI only
```

`RouteView.BuildLines` turns the model into the exact list of rows to draw, so the form never
tracks row indexes and the filter/sort behaviour can be tested without a window.

## Troubleshooting

**"Administrator rights are required to change firewall rules."** — Start the tool again with
*Run as administrator*.

**"Could not reach the Steam Web API."** — Network or TLS problem reaching
`api.steampowered.com`; check your connection and try **Change Game** to retry.

**"Steam does not have a relay config for app ID …"** — That app ID isn't valid. Check the
number in the game's store URL.

**Rules stay behind after a crash** — Launch the tool and press **Clear Rules**, or delete the
**SteamRouteTool** group in `wf.msc`.

## Credits

#### Froody
Tool creation.
#### Newtonsoft
Newtonsoft.Json package, licensed under MIT. https://github.com/JamesNK/Newtonsoft.Json/
#### Icon
https://github.com/feathericons/feather#feather
