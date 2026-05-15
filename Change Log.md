
## v6 — Current

- `IP:` prefix spacing corrected to align with `Status:` line; `AutoEllipsis` enabled to prevent fourth-row overflow
- `(+N)` count fixed to use distinct adapter aliases — dual-IP single NICs no longer inflate the count
- Primary IP sanitised to prevent array bleed into tile text; newline strip added in `Update-TileState`
- `[Set Static]` becomes `[Edit Static]` (amber, enabled) when server is already Static
- Adapter query broadened to collect all NICs with any IPv4 address, not just those with a gateway
- DNS fix — values trimmed before array build; `Preferred` = index 0, `Alternate` = index 1; errors now reported

---

## v5 — Async Set Static / Set DHCP

- Both operations moved off the UI thread into `Start-Job` — app stays interactive during changes
- Tile shows `Setting Static...` / `Setting DHCP...` immediately on submission
- Dedicated `ActionPollTimer` monitors action jobs; 40s watchdog kills stalled jobs
- Auto-rescan on success; error popup on failure with remote error message

---

## v4 — Colour, crash fixes, multi-IP tile

- IP type mismatch changed from red to orange; red now reserved for hard errors only
- `$clrTileBgOrange` colour definition added — its absence caused `BackColor null` crashes
- `$data.State` reference in `Set-IPTypeToggle` corrected to `$data.IpType`
- `servers.conf` detection hardened — `$PSCommandPath` and `Get-Location` added as fallbacks; resolved path shown in status bar
- Per-IP object emission in remote scriptblock fixes multi-IP wrapping on tiles
- Primary interface selected by DG subnet match rather than first gateway adapter
- `(+N)` suffix added to tile; detail output labels adapters as `[PRIMARY]`, `[Additional Adapter]`, `[Additional IP - same adapter]`

---

## v3 — Stability fixes

- `PSCustomObject` initialisers in `Add-DetailLine` and `Apply-JobResult` expanded to declare all nine properties upfront — fixes `Exception setting "StatusText"` and `BackColor` crashes
- `PrefixLength` ArrayList cast fixed — type checked before casting to `[int]`
- `$PSScriptRoot` null chain added: falls back to `$PSCommandPath` then `$MyInvocation.MyCommand.Path`

---

## v2 — Scan hang and timeout

- `New-PSSessionOption` added — `OpenTimeout 10s`, `OperationTimeout 20s`; dead servers fail in ~10s instead of hanging 3 minutes
- Per-job start timestamp tracked; poll timer kills jobs running over 35s with `Stop-Job`
- Timed-out servers show `Timed out` tile state and `[TIMEOUT]` detail line

---

## v1 — Initial release

- GUI tool to query and manage IPv4 config on remote Windows servers over WinRM
- Server input: one per line or comma-separated; auto-loads `servers.conf` from script directory
- `[Static]` / `[DHCP]` toggle colours tiles green (match) or red (mismatch); all queries run as background jobs
- Detail panel shows all adapters with IP, subnet, gateway, DNS
- **Set Static** — pre-filled dialog; applies `New-NetIPAddress` and `Set-DnsClientServerAddress`
- **Set DHCP** — confirmation required; enables DHCP and resets DNS
- **Rescan**, **Reboot** (with 30s auto-rescan countdown), **Clear**
- Status bar summary: `Done. Static: N DHCP: N Error: N`