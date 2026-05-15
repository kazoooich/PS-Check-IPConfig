#Requires -Version 5.1
<#
.SYNOPSIS
    Remote IP Address Checker & Configurator
.DESCRIPTION
    GUI tool to query and manage IP configuration on one or more remote Windows servers.
    Queries run as background jobs so the UI stays fully responsive.
    Click a tile to view detail. Toggle between Static/DHCP target state.
    Set Static IP or Set DHCP on selected server. Reboot available.
.NOTES
    Requirements:
      - PowerShell 5.1+
      - WinRM/PS Remoting enabled on target servers
      - Sufficient permissions on remote servers (local/domain admin)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region -- Colours -----------------------------------------------------------

$clrFormBg       = [System.Drawing.Color]::FromArgb( 30,  30,  46)
$clrPanelBg      = [System.Drawing.Color]::FromArgb( 36,  39,  58)
$clrOutputBg     = [System.Drawing.Color]::FromArgb( 24,  24,  37)
$clrText         = [System.Drawing.Color]::FromArgb(205, 214, 244)
$clrMuted        = [System.Drawing.Color]::FromArgb(166, 173, 200)
$clrBlue         = [System.Drawing.Color]::FromArgb(137, 180, 250)
$clrGreen        = [System.Drawing.Color]::FromArgb(166, 227, 161)
$clrOrange       = [System.Drawing.Color]::FromArgb(255, 165,   0)
$clrRed          = [System.Drawing.Color]::FromArgb(243, 139, 168)
$clrWarn         = [System.Drawing.Color]::FromArgb(250, 179, 135)
$clrAmber        = [System.Drawing.Color]::FromArgb(249, 226, 175)
$clrTileGrey     = [System.Drawing.Color]::FromArgb( 49,  50,  68)
$clrTileBgGrey   = [System.Drawing.Color]::FromArgb( 49,  50,  68)
$clrTileBgGreen  = [System.Drawing.Color]::FromArgb( 24,  60,  36)
$clrTileBgOrange = [System.Drawing.Color]::FromArgb( 70,  40,   0)
$clrTileBgRed    = [System.Drawing.Color]::FromArgb( 70,  20,  30)
$clrBtnActive    = [System.Drawing.Color]::FromArgb(137, 180, 250)
$clrBtnInactive  = [System.Drawing.Color]::FromArgb( 60,  64,  90)
$clrBtnDisabled  = [System.Drawing.Color]::FromArgb( 49,  50,  68)

#endregion

#region -- Layout constants --------------------------------------------------

$TILE_W       = 160
$TILE_H       = 72
$BTN_W        = 108
$STATUS_H     = 24
$DETAIL_H     = 160

#endregion

#region -- Build Form --------------------------------------------------------

$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'IP Address Checker'
$form.Size          = New-Object System.Drawing.Size(1100, 750)
$form.MinimumSize   = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Consolas', 9)
$form.BackColor     = $clrFormBg
$form.ForeColor     = $clrText

# -- Server input label -------------------------------------------------------
$lblServers           = New-Object System.Windows.Forms.Label
$lblServers.Text      = 'Servers (one per line or comma-separated):'
$lblServers.Location  = New-Object System.Drawing.Point(12, 12)
$lblServers.AutoSize  = $true
$lblServers.ForeColor = $clrBlue
$form.Controls.Add($lblServers)

# -- Server input textbox -----------------------------------------------------
$txtServers             = New-Object System.Windows.Forms.TextBox
$txtServers.Multiline   = $true
$txtServers.ScrollBars  = 'Vertical'
$txtServers.Location    = New-Object System.Drawing.Point(12, 32)
$txtServers.Size        = New-Object System.Drawing.Size(820, 72)
$txtServers.BackColor   = $clrPanelBg
$txtServers.ForeColor   = $clrGreen
$txtServers.BorderStyle = 'FixedSingle'
$txtServers.Anchor      = 'Top,Left,Right'
$form.Controls.Add($txtServers)

# -- Run button ---------------------------------------------------------------
$btnRun               = New-Object System.Windows.Forms.Button
$btnRun.Text          = '> Run'
$btnRun.Location      = New-Object System.Drawing.Point(844, 32)
$btnRun.Size          = New-Object System.Drawing.Size(120, 34)
$btnRun.BackColor     = $clrBlue
$btnRun.ForeColor     = $clrFormBg
$btnRun.FlatStyle     = 'Flat'
$btnRun.Font          = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
$btnRun.Anchor        = 'Top,Right'
$form.Controls.Add($btnRun)

# -- Clear button -------------------------------------------------------------
$btnClearTiles            = New-Object System.Windows.Forms.Button
$btnClearTiles.Text       = 'Clear'
$btnClearTiles.Location   = New-Object System.Drawing.Point(844, 70)
$btnClearTiles.Size       = New-Object System.Drawing.Size(120, 34)
$btnClearTiles.BackColor  = $clrTileGrey
$btnClearTiles.ForeColor  = $clrText
$btnClearTiles.FlatStyle  = 'Flat'
$btnClearTiles.Anchor     = 'Top,Right'
$form.Controls.Add($btnClearTiles)

# -- IP Type Required label ---------------------------------------------------
$lblIpType           = New-Object System.Windows.Forms.Label
$lblIpType.Text      = 'IP Type Required:'
$lblIpType.Location  = New-Object System.Drawing.Point(12, 114)
$lblIpType.AutoSize  = $true
$lblIpType.ForeColor = $clrBlue
$form.Controls.Add($lblIpType)

# -- Static toggle button -----------------------------------------------------
$btnToggleStatic              = New-Object System.Windows.Forms.Button
$btnToggleStatic.Text         = 'Static'
$btnToggleStatic.Location     = New-Object System.Drawing.Point(138, 109)
$btnToggleStatic.Size         = New-Object System.Drawing.Size(76, 26)
$btnToggleStatic.BackColor    = $clrBtnActive   # default selected
$btnToggleStatic.ForeColor    = $clrFormBg
$btnToggleStatic.FlatStyle    = 'Flat'
$btnToggleStatic.Font         = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnToggleStatic)

# -- DHCP toggle button -------------------------------------------------------
$btnToggleDHCP              = New-Object System.Windows.Forms.Button
$btnToggleDHCP.Text         = 'DHCP'
$btnToggleDHCP.Location     = New-Object System.Drawing.Point(218, 109)
$btnToggleDHCP.Size         = New-Object System.Drawing.Size(76, 26)
$btnToggleDHCP.BackColor    = $clrBtnInactive
$btnToggleDHCP.ForeColor    = $clrText
$btnToggleDHCP.FlatStyle    = 'Flat'
$btnToggleDHCP.Font         = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($btnToggleDHCP)

# -- Server Status label ------------------------------------------------------
$lblTiles           = New-Object System.Windows.Forms.Label
$lblTiles.Text      = 'Server Status:'
$lblTiles.Location  = New-Object System.Drawing.Point(12, 143)
$lblTiles.AutoSize  = $true
$lblTiles.ForeColor = $clrBlue
$form.Controls.Add($lblTiles)

# -- Tile FlowLayoutPanel -----------------------------------------------------
$pnlTiles               = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlTiles.Location      = New-Object System.Drawing.Point(12, 163)
$pnlTiles.AutoScroll    = $true
$pnlTiles.BackColor     = $clrFormBg
$pnlTiles.BorderStyle   = 'None'
$pnlTiles.FlowDirection = 'LeftToRight'
$pnlTiles.WrapContents  = $true
$pnlTiles.Anchor        = 'Top,Left,Right,Bottom'
$form.Controls.Add($pnlTiles)

# -- Detail label -------------------------------------------------------------
$lblDetail           = New-Object System.Windows.Forms.Label
$lblDetail.Text      = 'IP Detail:  (click a tile above to view)'
$lblDetail.AutoSize  = $true
$lblDetail.ForeColor = $clrBlue
$lblDetail.Anchor    = 'Bottom,Left'
$form.Controls.Add($lblDetail)

# -- Detail RichTextBox -------------------------------------------------------
$rtbDetail             = New-Object System.Windows.Forms.RichTextBox
$rtbDetail.ReadOnly    = $true
$rtbDetail.ScrollBars  = 'Vertical'
$rtbDetail.BackColor   = $clrOutputBg
$rtbDetail.ForeColor   = $clrText
$rtbDetail.BorderStyle = 'FixedSingle'
$rtbDetail.Font        = New-Object System.Drawing.Font('Consolas', 9)
$rtbDetail.WordWrap    = $true
$rtbDetail.Anchor      = 'Bottom,Left,Right'
$form.Controls.Add($rtbDetail)

# -- Action buttons -----------------------------------------------------------
$btnRescan            = New-Object System.Windows.Forms.Button
$btnRescan.Text       = 'Rescan'
$btnRescan.Size       = New-Object System.Drawing.Size($BTN_W, 30)
$btnRescan.BackColor  = $clrBlue
$btnRescan.ForeColor  = $clrFormBg
$btnRescan.FlatStyle  = 'Flat'
$btnRescan.Font       = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
$btnRescan.Enabled    = $false
$btnRescan.Anchor     = 'Bottom,Right'
$form.Controls.Add($btnRescan)

$btnSetStatic            = New-Object System.Windows.Forms.Button
$btnSetStatic.Text       = 'Set Static'
$btnSetStatic.Size       = New-Object System.Drawing.Size($BTN_W, 30)
$btnSetStatic.BackColor  = $clrBtnDisabled
$btnSetStatic.ForeColor  = $clrMuted
$btnSetStatic.FlatStyle  = 'Flat'
$btnSetStatic.Font       = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
$btnSetStatic.Enabled    = $false
$btnSetStatic.Anchor     = 'Bottom,Right'
$form.Controls.Add($btnSetStatic)

$btnSetDHCP            = New-Object System.Windows.Forms.Button
$btnSetDHCP.Text       = 'Set DHCP'
$btnSetDHCP.Size       = New-Object System.Drawing.Size($BTN_W, 30)
$btnSetDHCP.BackColor  = $clrBtnDisabled
$btnSetDHCP.ForeColor  = $clrMuted
$btnSetDHCP.FlatStyle  = 'Flat'
$btnSetDHCP.Font       = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
$btnSetDHCP.Enabled    = $false
$btnSetDHCP.Anchor     = 'Bottom,Right'
$form.Controls.Add($btnSetDHCP)

$btnReboot            = New-Object System.Windows.Forms.Button
$btnReboot.Text       = 'Reboot'
$btnReboot.Size       = New-Object System.Drawing.Size($BTN_W, 30)
$btnReboot.BackColor  = $clrRed
$btnReboot.ForeColor  = $clrFormBg
$btnReboot.FlatStyle  = 'Flat'
$btnReboot.Font       = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
$btnReboot.Enabled    = $false
$btnReboot.Anchor     = 'Bottom,Right'
$form.Controls.Add($btnReboot)

# -- Status bar ---------------------------------------------------------------
$statusBar   = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Ready.'
$statusBar.Items.Add($statusLabel) | Out-Null
$statusBar.BackColor = $clrOutputBg
$statusBar.ForeColor = $clrGreen
$form.Controls.Add($statusBar)

#endregion

#region -- Layout / resize ---------------------------------------------------

function Update-Layout {
    $w = $form.ClientSize.Width
    $h = $form.ClientSize.Height

    $txtServers.Size        = New-Object System.Drawing.Size(($w - 180), 72)
    $btnRun.Location        = New-Object System.Drawing.Point(($w - 132), 32)
    $btnClearTiles.Location = New-Object System.Drawing.Point(($w - 132), 70)

    $detailH   = $DETAIL_H
    $detailBot = $h - $STATUS_H - 8
    $detailTop = $detailBot - $detailH
    $lblTop    = $detailTop - 18

    $detailW   = $w - 24 - $BTN_W - 16
    $btnX      = $w - $BTN_W - 12

    $lblDetail.Location  = New-Object System.Drawing.Point(12, $lblTop)
    $rtbDetail.Location  = New-Object System.Drawing.Point(12, $detailTop)
    $rtbDetail.Size      = New-Object System.Drawing.Size($detailW, $detailH)

    $btnRescan.Location    = New-Object System.Drawing.Point($btnX, $detailTop)
    $btnSetStatic.Location = New-Object System.Drawing.Point($btnX, ($detailTop + 38))
    $btnSetDHCP.Location   = New-Object System.Drawing.Point($btnX, ($detailTop + 76))
    $btnReboot.Location    = New-Object System.Drawing.Point($btnX, ($detailTop + 114))

    $tilePanelTop = 163
    $tilePanelH   = [Math]::Max(60, $lblTop - $tilePanelTop - 4)
    $pnlTiles.Location = New-Object System.Drawing.Point(12, $tilePanelTop)
    $pnlTiles.Size     = New-Object System.Drawing.Size(($w - 24), $tilePanelH)
}

$form.Add_Resize({ Update-Layout })
Update-Layout

#endregion

#region -- State store -------------------------------------------------------

$script:ServerData      = @{}
$script:SelectedServer  = $null
$script:ActiveJobs      = @{}
$script:JobStartTime    = @{}
$script:PendingCount    = 0
$script:RebootTimer     = $null
$script:RequiredIPType  = 'Static'   # default
$script:JobTimeoutSecs  = 35         # kill any query job still running after this long
$script:ActionJobs      = @{}        # background Set Static / Set DHCP jobs keyed by server

#endregion

#region -- Toggle button logic -----------------------------------------------

function Set-IPTypeToggle {
    param([string]$Type)
    $script:RequiredIPType = $Type

    if ($Type -eq 'Static') {
        $btnToggleStatic.BackColor = $clrBtnActive
        $btnToggleStatic.ForeColor = $clrFormBg
        $btnToggleStatic.Font      = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
        $btnToggleDHCP.BackColor   = $clrBtnInactive
        $btnToggleDHCP.ForeColor   = $clrText
        $btnToggleDHCP.Font        = New-Object System.Drawing.Font('Consolas', 9)
    } else {
        $btnToggleDHCP.BackColor   = $clrBtnActive
        $btnToggleDHCP.ForeColor   = $clrFormBg
        $btnToggleDHCP.Font        = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
        $btnToggleStatic.BackColor = $clrBtnInactive
        $btnToggleStatic.ForeColor = $clrText
        $btnToggleStatic.Font      = New-Object System.Drawing.Font('Consolas', 9)
    }

    # Refresh all tile colours based on new required type
    foreach ($tile in $pnlTiles.Controls) {
        $srv = $tile.Tag
        if ($script:ServerData.ContainsKey($srv)) {
            $data = $script:ServerData[$srv]
            if ($data.IpType -ne 'pending') {
                Update-TileState $srv $data.IpType $data.StatusText $data.IpText
            }
        }
    }
}

$btnToggleStatic.Add_Click({ Set-IPTypeToggle 'Static' })
$btnToggleDHCP.Add_Click({   Set-IPTypeToggle 'DHCP'   })

#endregion

#region -- Tile helpers ------------------------------------------------------

function Get-TileColors {
    param([string]$IpType)
    if ($IpType -eq 'pending') { return @{ Bg = $clrTileBgGrey;   Text = $clrMuted  } }
    if ($IpType -eq 'error')   { return @{ Bg = $clrTileBgRed;    Text = $clrRed    } }
    if ($IpType -eq $script:RequiredIPType) {
        return @{ Bg = $clrTileBgGreen;  Text = $clrGreen  }
    } else {
        return @{ Bg = $clrTileBgOrange; Text = $clrOrange }
    }
}

function New-ServerTile {
    param([string]$ServerName)

    $tile             = New-Object System.Windows.Forms.Panel
    $tile.Size        = New-Object System.Drawing.Size($TILE_W, $TILE_H)
    $tile.BackColor   = $clrTileBgGrey
    $tile.BorderStyle = 'FixedSingle'
    $tile.Margin      = New-Object System.Windows.Forms.Padding(4)
    $tile.Cursor      = [System.Windows.Forms.Cursors]::Hand
    $tile.Tag         = $ServerName

    $lblName              = New-Object System.Windows.Forms.Label
    $lblName.Text         = $ServerName.ToUpper()
    $lblName.Font         = New-Object System.Drawing.Font('Consolas', 8, [System.Drawing.FontStyle]::Bold)
    $lblName.ForeColor    = $clrBlue
    $lblName.BackColor    = [System.Drawing.Color]::Transparent
    $lblName.Location     = New-Object System.Drawing.Point(6, 5)
    $lblName.Size         = New-Object System.Drawing.Size(($TILE_W - 10), 16)
    $lblName.AutoEllipsis = $true
    $tile.Controls.Add($lblName)

    $lblStatus            = New-Object System.Windows.Forms.Label
    $lblStatus.Text       = 'Pending...'
    $lblStatus.Font       = New-Object System.Drawing.Font('Consolas', 8)
    $lblStatus.ForeColor  = $clrMuted
    $lblStatus.BackColor  = [System.Drawing.Color]::Transparent
    $lblStatus.Location   = New-Object System.Drawing.Point(6, 24)
    $lblStatus.Size       = New-Object System.Drawing.Size(($TILE_W - 10), 16)
    $lblStatus.Tag        = 'status'
    $tile.Controls.Add($lblStatus)

    $lblIp                = New-Object System.Windows.Forms.Label
    $lblIp.Text           = ''
    $lblIp.Font           = New-Object System.Drawing.Font('Consolas', 8)
    $lblIp.ForeColor      = $clrMuted
    $lblIp.BackColor      = [System.Drawing.Color]::Transparent
    $lblIp.Location       = New-Object System.Drawing.Point(6, 43)
    $lblIp.Size           = New-Object System.Drawing.Size(($TILE_W - 10), 22)
    $lblIp.Tag            = 'ip'
    $tile.Controls.Add($lblIp)

    $clickHandler = {
        param($s, $e)
        $tp = if ($s -is [System.Windows.Forms.Panel]) { $s } else { $s.Parent }
        Select-Tile -ServerName $tp.Tag
    }
    $tile.Add_Click($clickHandler)
    foreach ($c in $tile.Controls) { $c.Add_Click($clickHandler) }

    return $tile
}

function Update-TileState {
    param(
        [string]$ServerName,
        [string]$IpType,        # 'Static','DHCP','pending','error'
        [string]$StatusText,
        [string]$IpText
    )
    $tile = $pnlTiles.Controls | Where-Object { $_.Tag -eq $ServerName } | Select-Object -First 1
    if (-not $tile) { return }

    $colors = Get-TileColors -IpType $IpType
    $tile.BackColor = $colors.Bg

    foreach ($c in $tile.Controls) {
        switch ($c.Tag) {
            'status' { $c.Text = $StatusText; $c.ForeColor = $colors.Text }
            'ip'     { $c.Text = $IpText;     $c.ForeColor = $colors.Text }
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Select-Tile {
    param([string]$ServerName)
    $script:SelectedServer = $ServerName

    foreach ($tile in $pnlTiles.Controls) {
        $tile.BorderStyle = if ($tile.Tag -eq $ServerName) { 'Fixed3D' } else { 'FixedSingle' }
    }

    $rtbDetail.Clear()
    $lblDetail.Text = "IP Detail:  $($ServerName.ToUpper())"

    if ($script:ServerData.ContainsKey($ServerName)) {
        $data = $script:ServerData[$ServerName]
        foreach ($entry in $data.Lines) {
            $rtbDetail.SelectionStart  = $rtbDetail.TextLength
            $rtbDetail.SelectionLength = 0
            $rtbDetail.SelectionColor  = $entry.Color
            $rtbDetail.AppendText("$($entry.Text)`n")
        }
        $rtbDetail.ScrollToCaret()

        # Update action button states based on current server IP type
        $btnRescan.Enabled = $true
        $btnReboot.Enabled = $data.Success

        if ($data.Success) {
            # Set Static: grey out if already Static
            if ($data.IpType -eq 'Static') {
                $btnSetStatic.Enabled   = $false
                $btnSetStatic.BackColor = $clrBtnDisabled
                $btnSetStatic.ForeColor = $clrMuted
            } else {
                $btnSetStatic.Enabled   = $true
                $btnSetStatic.BackColor = $clrGreen
                $btnSetStatic.ForeColor = $clrFormBg
            }

            # Set DHCP: grey out if already DHCP
            if ($data.IpType -eq 'DHCP') {
                $btnSetDHCP.Enabled   = $false
                $btnSetDHCP.BackColor = $clrBtnDisabled
                $btnSetDHCP.ForeColor = $clrMuted
            } else {
                $btnSetDHCP.Enabled   = $true
                $btnSetDHCP.BackColor = $clrWarn
                $btnSetDHCP.ForeColor = $clrFormBg
            }
        } else {
            $btnSetStatic.Enabled   = $false
            $btnSetStatic.BackColor = $clrBtnDisabled
            $btnSetStatic.ForeColor = $clrMuted
            $btnSetDHCP.Enabled     = $false
            $btnSetDHCP.BackColor   = $clrBtnDisabled
            $btnSetDHCP.ForeColor   = $clrMuted
        }
    } else {
        $btnRescan.Enabled = $true
        $btnReboot.Enabled = $false
        $btnSetStatic.Enabled   = $false
        $btnSetStatic.BackColor = $clrBtnDisabled
        $btnSetStatic.ForeColor = $clrMuted
        $btnSetDHCP.Enabled     = $false
        $btnSetDHCP.BackColor   = $clrBtnDisabled
        $btnSetDHCP.ForeColor   = $clrMuted

        $rtbDetail.SelectionColor = $clrMuted
        $rtbDetail.AppendText('No data yet. Press Run or Rescan.')
    }

    $statusLabel.Text = "Selected: $ServerName"
}

#endregion

#region -- Detail line store -------------------------------------------------

function Add-DetailLine {
    param(
        [string]$ServerName,
        [string]$Text,
        [System.Drawing.Color]$Color
    )
    if (-not $script:ServerData.ContainsKey($ServerName)) {
        $script:ServerData[$ServerName] = [PSCustomObject]@{
            Lines       = [System.Collections.Generic.List[PSCustomObject]]::new()
            IpType      = 'pending'
            StatusText  = ''
            IpText      = ''
            PrimaryIP   = ''
            PrimaryMask = ''
            PrimaryGW   = ''
            PrimaryDNS1 = ''
            PrimaryDNS2 = ''
            Success     = $false
        }
    }
    $script:ServerData[$ServerName].Lines.Add([PSCustomObject]@{ Text = $Text; Color = $Color })

    if ($script:SelectedServer -eq $ServerName) {
        $rtbDetail.SelectionStart  = $rtbDetail.TextLength
        $rtbDetail.SelectionLength = 0
        $rtbDetail.SelectionColor  = $Color
        $rtbDetail.AppendText("$Text`n")
        $rtbDetail.ScrollToCaret()
    }
}

#endregion

#region -- Colour resolver ---------------------------------------------------

function Resolve-Color {
    param([string]$Name)
    switch ($Name) {
        'Green'  { return $clrGreen  }
        'Red'    { return $clrRed    }
        'Orange' { return $clrOrange }
        'Warn'   { return $clrWarn   }
        'Amber'  { return $clrAmber  }
        'Blue'   { return $clrBlue   }
        'Muted'  { return $clrMuted  }
        default  { return $clrMuted  }
    }
}

#endregion

#region -- Background job scriptblock ----------------------------------------

$script:QueryScriptBlock = {
    param([string]$Server)

    $out = [System.Collections.Generic.List[PSCustomObject]]::new()
    function Line([string]$t, [string]$c) {
        $out.Add([PSCustomObject]@{ Text = $t; ColorName = $c })
    }

    Line "== $($Server.ToUpper()) ==" 'Amber'
    Line '' 'Muted'

    try {
        $sessionOpt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 20000 -CancelTimeout 5000
        $result = Invoke-Command -ComputerName $Server -SessionOption $sessionOpt -ScriptBlock {
            $adapters = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null }

            if (-not $adapters) {
                # Fall back to any adapter with an IP
                $adapters = Get-NetIPConfiguration | Where-Object { $_.IPv4Address -ne $null }
            }

            $adapterList = foreach ($a in $adapters) {
                $dhcpEnabled = $false
                try {
                    $netAdapter = Get-NetIPInterface -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction Stop
                    $dhcpEnabled = ($netAdapter.Dhcp -eq 'Enabled')
                } catch {}

                $dnsServers = @()
                try {
                    $dns = Get-DnsClientServerAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
                    if ($dns) { $dnsServers = $dns.ServerAddresses }
                } catch {}

                $gw = if ($a.IPv4DefaultGateway) { $a.IPv4DefaultGateway.NextHop } else { '' }

                # IPv4Address can be an array if the adapter has multiple IPs - emit one object per IP
                $ipEntries = @($a.IPv4Address)
                if (-not $ipEntries -or $ipEntries.Count -eq 0) {
                    # No IPs on this adapter - emit a placeholder so adapter still appears
                    [PSCustomObject]@{
                        Alias       = $a.InterfaceAlias
                        Index       = $a.InterfaceIndex
                        IPAddress   = ''
                        PrefixLen   = 0
                        Gateway     = $gw
                        DhcpEnabled = $dhcpEnabled
                        DNS         = $dnsServers
                    }
                } else {
                    $firstOnAdapter = $true
                    foreach ($ipEntry in $ipEntries) {
                        $rawPrefix = $ipEntry.PrefixLength
                        $pfx = if ($rawPrefix -is [System.Collections.IEnumerable] -and $rawPrefix -isnot [string]) {
                            [int]($rawPrefix | Select-Object -First 1)
                        } else { [int]$rawPrefix }

                        [PSCustomObject]@{
                            Alias       = $a.InterfaceAlias
                            Index       = $a.InterfaceIndex
                            IPAddress   = [string]$ipEntry.IPAddress
                            PrefixLen   = $pfx
                            # Only carry GW/DNS on first IP entry per adapter
                            Gateway     = if ($firstOnAdapter) { $gw } else { '' }
                            DhcpEnabled = $dhcpEnabled
                            DNS         = if ($firstOnAdapter) { $dnsServers } else { @() }
                        }
                        $firstOnAdapter = $false
                    }
                }
            }
            @($adapterList)
        } -ErrorAction Stop

        if (-not $result) {
            Line '  [INFO] No network adapters with IP configuration found.' 'Warn'
            return [PSCustomObject]@{
                Server      = $Server
                Lines       = $out
                IpType      = 'error'
                StatusText  = 'No adapters'
                IpText      = ''
                PrimaryIP   = ''
                PrimaryMask = ''
                PrimaryGW   = ''
                PrimaryDNS1 = ''
                PrimaryDNS2 = ''
                Success     = $false
            }
        }

        # Convert prefix length to subnet mask
        function PrefixToMask([int]$prefix) {
            if ($prefix -lt 0 -or $prefix -gt 32) { return '0.0.0.0' }
            $mask = [uint32]([Math]::Pow(2,32) - [Math]::Pow(2, 32 - $prefix))
            return "$( ($mask -shr 24) -band 255 ).$( ($mask -shr 16) -band 255 ).$( ($mask -shr 8) -band 255 ).$( $mask -band 255 )"
        }

        # Helper: check if an IP is in the same subnet as the gateway
        function Test-SameSubnet([string]$ip, [string]$gw, [int]$prefix) {
            try {
                $shift   = 32 - $prefix
                $ipInt   = [uint32]([System.Net.IPAddress]::Parse($ip).GetAddressBytes() |
                               ForEach-Object { $_ } |
                               ForEach-Object -Begin { $acc = [uint32]0 } `
                                              -Process { $acc = ($acc -shl 8) -bor $_ } `
                                              -End { $acc })
                $gwInt   = [uint32]([System.Net.IPAddress]::Parse($gw).GetAddressBytes() |
                               ForEach-Object { $_ } |
                               ForEach-Object -Begin { $acc = [uint32]0 } `
                                              -Process { $acc = ($acc -shl 8) -bor $_ } `
                                              -End { $acc })
                $mask    = if ($shift -ge 32) { [uint32]0 } else { [uint32](0xFFFFFFFF -shl $shift) }
                return (($ipInt -band $mask) -eq ($gwInt -band $mask))
            } catch { return $false }
        }

        # Sort adapters: gateway-bearing first, then others
        $sorted = @($result | Sort-Object { if ($_.Gateway) { 0 } else { 1 } })

        # Pick the primary adapter: the one whose IP is in the same subnet as the DG.
        # Fall back to first gateway-bearing adapter if none match subnet.
        $primary = $null
        foreach ($a in $sorted) {
            if ($a.Gateway -and $a.IPAddress -and (Test-SameSubnet $a.IPAddress $a.Gateway $a.PrefixLen)) {
                $primary = $a
                break
            }
        }
        if (-not $primary) { $primary = $sorted[0] }

        # Re-order so primary is always first in output
        $ordered = @($primary) + @($sorted | Where-Object { $_ -ne $primary })

        $primaryIpType  = if ($primary.DhcpEnabled) { 'DHCP' } else { 'Static' }
        $primaryIP      = $primary.IPAddress
        $primaryGW      = $primary.Gateway
        $primaryDns1    = if ($primary.DNS.Count -gt 0) { $primary.DNS[0] } else { '' }
        $primaryDns2    = if ($primary.DNS.Count -gt 1) { $primary.DNS[1] } else { '' }
        $primaryMask    = PrefixToMask $primary.PrefixLen

        $extraCount     = $ordered.Count - 1
        $extraSuffix    = if ($extraCount -gt 0) { " (+$extraCount)" } else { '' }
        $tileStatusText = "Status: $primaryIpType"
        $tileIpText     = "IP:     $primaryIP$extraSuffix"

        $prevAlias = $null
        foreach ($a in $ordered) {
            $ipType    = if ($a.DhcpEnabled) { 'DHCP' } else { 'Static' }
            $ipColor   = if ($a.DhcpEnabled) { 'Blue' } else { 'Green' }
            $isPrimary = ($a -eq $primary)
            $sameAdapter = ($a.Alias -eq $prevAlias)

            if ($isPrimary) {
                $tag = '  [PRIMARY - DG Subnet Match]'
            } elseif ($sameAdapter) {
                $tag = '  [Additional IP - same adapter]'
            } else {
                $tag = '  [Additional Adapter]'
            }

            # Only print adapter name header when it changes
            if (-not $sameAdapter) {
                Line "  Adapter : $($a.Alias)"                 'Muted'
                Line "  Type    : $ipType"                     $ipColor
            }
            Line "  IP Addr : $($a.IPAddress)/$($a.PrefixLen)" 'Muted'
            Line "  Subnet  : $(PrefixToMask $a.PrefixLen)"    'Muted'
            if ($a.Gateway) {
                Line "  Gateway : $($a.Gateway)"               'Muted'
            }
            if ($a.DNS.Count -gt 0) {
                Line "  DNS 1   : $($a.DNS[0])"                'Muted'
            }
            if ($a.DNS.Count -gt 1) {
                Line "  DNS 2   : $($a.DNS[1])"                'Muted'
            }
            $prevAlias = $a.Alias
            Line $tag                                           'Amber'
            Line '' 'Muted'
        }

        return [PSCustomObject]@{
            Server      = $Server
            Lines       = $out
            IpType      = $primaryIpType
            StatusText  = $tileStatusText
            IpText      = $tileIpText
            PrimaryIP   = $primaryIP
            PrimaryMask = $primaryMask
            PrimaryGW   = $primaryGW
            PrimaryDNS1 = $primaryDns1
            PrimaryDNS2 = $primaryDns2
            Success     = $true
        }

    } catch {
        $msg = $_.Exception.Message
        Line "  [ERROR] $msg" 'Red'
        return [PSCustomObject]@{
            Server      = $Server
            Lines       = $out
            IpType      = 'error'
            StatusText  = 'Connection failed'
            IpText      = ''
            PrimaryIP   = ''
            PrimaryMask = ''
            PrimaryGW   = ''
            PrimaryDNS1 = ''
            PrimaryDNS2 = ''
            Success     = $false
            Error       = $msg
        }
    }
}

#endregion

#region -- Apply job result to UI --------------------------------------------

function Apply-JobResult {
    param($result)
    $srv = $result.Server

    if ($script:ServerData.ContainsKey($srv)) {
        $script:ServerData[$srv].Lines.Clear()
    } else {
        $script:ServerData[$srv] = [PSCustomObject]@{
            Lines       = [System.Collections.Generic.List[PSCustomObject]]::new()
            IpType      = 'pending'
            StatusText  = ''
            IpText      = ''
            PrimaryIP   = ''
            PrimaryMask = ''
            PrimaryGW   = ''
            PrimaryDNS1 = ''
            PrimaryDNS2 = ''
            Success     = $false
        }
    }
    if ($script:SelectedServer -eq $srv) { $rtbDetail.Clear() }

    foreach ($entry in $result.Lines) {
        Add-DetailLine $srv $entry.Text (Resolve-Color $entry.ColorName)
    }

    # Store metadata
    $script:ServerData[$srv].IpType      = $result.IpType
    $script:ServerData[$srv].StatusText  = $result.StatusText
    $script:ServerData[$srv].IpText      = $result.IpText
    $script:ServerData[$srv].PrimaryIP   = $result.PrimaryIP
    $script:ServerData[$srv].PrimaryMask = $result.PrimaryMask
    $script:ServerData[$srv].PrimaryGW   = $result.PrimaryGW
    $script:ServerData[$srv].PrimaryDNS1 = $result.PrimaryDNS1
    $script:ServerData[$srv].PrimaryDNS2 = $result.PrimaryDNS2
    $script:ServerData[$srv].Success     = $result.Success

    Update-TileState $srv $result.IpType $result.StatusText $result.IpText

    if ($script:SelectedServer -eq $srv) {
        Select-Tile $srv
    }
}

#endregion

#region -- Action job polling timer ------------------------------------------
# Polls background Set Static / Set DHCP jobs separately from query jobs
# so they never interfere with each other

$script:ActionPollTimer          = New-Object System.Windows.Forms.Timer
$script:ActionPollTimer.Interval = 500

$script:ActionPollTimer.Add_Tick({
    if ($script:ActionJobs.Count -eq 0) { $script:ActionPollTimer.Stop(); return }

    $completed = @()
    $now = [datetime]::Now
    foreach ($kv in $script:ActionJobs.GetEnumerator()) {
        $srv     = $kv.Key
        $jobInfo = $kv.Value
        if ($jobInfo.Job.State -in 'Completed','Failed','Stopped') {
            $completed += $srv
        } elseif (($now - $jobInfo.StartTime).TotalSeconds -gt 40) {
            Stop-Job -Job $jobInfo.Job -ErrorAction SilentlyContinue
            $completed += $srv
        }
    }

    foreach ($srv in $completed) {
        $jobInfo = $script:ActionJobs[$srv]
        $job     = $jobInfo.Job
        try {
            if ($job.State -eq 'Stopped') {
                Add-DetailLine $srv "  [TIMEOUT] Action timed out after 40s" $clrWarn
                Update-TileState $srv 'error' 'Action timed out' ''
                $statusLabel.Text = "Action timed out on $srv"
            } else {
                $result = Receive-Job -Job $job -ErrorAction Stop
                if ($result -and $result.Success) {
                    if ($result.Action -eq 'SetStatic') {
                        Add-DetailLine $srv "  Set Static succeeded. Rescanning..." $clrGreen
                        $statusLabel.Text = "Static IP set on $srv. Rescanning..."
                    } else {
                        Add-DetailLine $srv "  Set DHCP succeeded. Rescanning..." $clrGreen
                        $statusLabel.Text = "DHCP enabled on $srv. Rescanning..."
                    }
                    # Auto-rescan now the change is done
                    Start-ServerQuery -Servers @($srv) -IsRescan $true
                } elseif ($result) {
                    $errMsg = $result.Error
                    Add-DetailLine $srv "  [ERROR] $errMsg" $clrRed
                    Update-TileState $srv 'error' 'Action failed' ''
                    $statusLabel.Text = "Action failed on $srv"
                    [System.Windows.Forms.MessageBox]::Show(
                        "Action failed on $srv .`n`n$errMsg",
                        'Action Failed',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                    # Re-enable buttons since it failed
                    if ($script:SelectedServer -eq $srv -and $script:ServerData.ContainsKey($srv)) {
                        Select-Tile $srv
                    }
                }
            }
        } catch {
            Add-DetailLine $srv "  [ERROR] $($_.Exception.Message)" $clrRed
            Update-TileState $srv 'error' 'Action error' ''
        } finally {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            $script:ActionJobs.Remove($srv)
        }
    }

    if ($script:ActionJobs.Count -eq 0) { $script:ActionPollTimer.Stop() }
})

#endregion

#region -- Job polling timer -------------------------------------------------

$script:PollTimer          = New-Object System.Windows.Forms.Timer
$script:PollTimer.Interval = 500

$script:PollTimer.Add_Tick({
    if ($script:ActiveJobs.Count -eq 0) { return }

    $completed = @()
    $now = [datetime]::Now
    foreach ($kv in $script:ActiveJobs.GetEnumerator()) {
        $srv = $kv.Key
        if ($kv.Value.State -in 'Completed','Failed','Stopped') {
            $completed += $srv
        } elseif ($script:JobStartTime.ContainsKey($srv)) {
            $elapsed = ($now - $script:JobStartTime[$srv]).TotalSeconds
            if ($elapsed -gt $script:JobTimeoutSecs) {
                # Job is hung - stop it and mark as timed out
                Stop-Job -Job $kv.Value -ErrorAction SilentlyContinue
                $completed += $srv
            }
        }
    }

    foreach ($srv in $completed) {
        $job = $script:ActiveJobs[$srv]
        try {
            if ($job.State -eq 'Stopped') {
                # Job was killed by watchdog - timed out
                Add-DetailLine $srv "  [TIMEOUT] No response after $($script:JobTimeoutSecs)s" $clrWarn
                Update-TileState $srv 'error' 'Timed out' ''
            } else {
                $result = Receive-Job -Job $job -ErrorAction Stop
                if ($result) { Apply-JobResult $result }
                else         { Update-TileState $srv 'error' 'No data returned' '' }
            }
        } catch {
            Add-DetailLine $srv "  [ERROR] $($_.Exception.Message)" $clrRed
            Update-TileState $srv 'error' 'Job error' ''
        } finally {
            Remove-Job -Job $job -Force
            $script:ActiveJobs.Remove($srv)
            $script:JobStartTime.Remove($srv)
            $script:PendingCount--
        }
    }

    if ($script:PendingCount -gt 0) {
        $statusLabel.Text = "Querying... $($script:PendingCount) server(s) remaining"
    } else {
        $static = ($script:ServerData.Values | Where-Object { $_.IpType -eq 'Static' }).Count
        $dhcp   = ($script:ServerData.Values | Where-Object { $_.IpType -eq 'DHCP'   }).Count
        $err    = ($script:ServerData.Values | Where-Object { $_.IpType -eq 'error'  }).Count
        $statusLabel.Text = "Done.  Static: $static  DHCP: $dhcp  Error: $err"
        $script:PollTimer.Stop()
        $btnRun.Enabled = $true
        $btnRun.Text    = '> Run'
    }
})

#endregion

#region -- Post-reboot countdown ---------------------------------------------

function Start-RebootRescanCountdown {
    param([string]$ServerName)
    if ($script:RebootTimer) {
        $script:RebootTimer.Stop()
        $script:RebootTimer.Dispose()
        $script:RebootTimer = $null
    }
    $script:RebootRescanServer = $ServerName
    $script:RebootCountdown    = 30
    $script:RebootTimer          = New-Object System.Windows.Forms.Timer
    $script:RebootTimer.Interval = 1000
    $script:RebootTimer.Add_Tick({
        $script:RebootCountdown--
        $statusLabel.Text = "Auto-rescan $($script:RebootRescanServer) in $($script:RebootCountdown)s..."
        if ($script:RebootCountdown -le 0) {
            $script:RebootTimer.Stop()
            $script:RebootTimer.Dispose()
            $script:RebootTimer = $null
            $statusLabel.Text   = "Auto-rescanning $($script:RebootRescanServer)..."
            Start-ServerQuery -Servers @($script:RebootRescanServer) -IsRescan $true
        }
    })
    $script:RebootTimer.Start()
}

#endregion

#region -- Launch queries ----------------------------------------------------

function Start-ServerQuery {
    param([string[]]$Servers, [bool]$IsRescan = $false)

    $script:PollTimer.Stop()
    foreach ($j in $script:ActiveJobs.Values) { Remove-Job -Job $j -Force -ErrorAction SilentlyContinue }
    $script:ActiveJobs = @{}

    if (-not $IsRescan) {
        $script:ServerData     = @{}
        $script:JobStartTime   = @{}
        $script:SelectedServer = $null
        $btnRescan.Enabled     = $false
        $btnSetStatic.Enabled  = $false
        $btnSetStatic.BackColor = $clrBtnDisabled
        $btnSetStatic.ForeColor = $clrMuted
        $btnSetDHCP.Enabled    = $false
        $btnSetDHCP.BackColor  = $clrBtnDisabled
        $btnSetDHCP.ForeColor  = $clrMuted
        $btnReboot.Enabled     = $false
        $rtbDetail.Clear()
        $lblDetail.Text        = 'IP Detail:  (click a tile above to view)'
        $pnlTiles.Controls.Clear()

        foreach ($srv in $Servers) {
            $tile = New-ServerTile -ServerName $srv
            $pnlTiles.Controls.Add($tile)
        }
    } else {
        $srv = $Servers[0]
        if ($script:ServerData.ContainsKey($srv)) { $script:ServerData[$srv].Lines.Clear() }
        Update-TileState $srv 'pending' 'Rescanning...' ''
        if ($script:SelectedServer -eq $srv) { $rtbDetail.Clear() }
        $btnRescan.Enabled    = $false
        $btnSetStatic.Enabled = $false
        $btnSetStatic.BackColor = $clrBtnDisabled
        $btnSetStatic.ForeColor = $clrMuted
        $btnSetDHCP.Enabled   = $false
        $btnSetDHCP.BackColor  = $clrBtnDisabled
        $btnSetDHCP.ForeColor  = $clrMuted
        $btnReboot.Enabled    = $false
    }

    [System.Windows.Forms.Application]::DoEvents()

    $script:PendingCount = $Servers.Count
    $btnRun.Enabled      = $false
    $btnRun.Text         = 'Running...'

    foreach ($srv in $Servers) {
        $job = Start-Job -ScriptBlock $script:QueryScriptBlock -ArgumentList $srv
        $script:ActiveJobs[$srv]         = $job
        $script:JobStartTime[$srv]       = [datetime]::Now
    }

    $script:PollTimer.Start()
    $statusLabel.Text = "Querying $($Servers.Count) server(s) in background..."
}

#endregion

#region -- Run button --------------------------------------------------------

$btnRun.Add_Click({
    $rawInput = $txtServers.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($rawInput)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please enter at least one server name.',
            'No servers specified',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }
    $servers = $rawInput -split '[\r\n,]+' |
               ForEach-Object { $_.Trim() } |
               Where-Object   { $_ -ne '' } |
               Select-Object  -Unique
    Start-ServerQuery -Servers $servers -IsRescan $false
})

#endregion

#region -- Rescan button -----------------------------------------------------

$btnRescan.Add_Click({
    if (-not $script:SelectedServer) { return }
    if ($script:RebootTimer) {
        $script:RebootTimer.Stop()
        $script:RebootTimer.Dispose()
        $script:RebootTimer = $null
    }
    Start-ServerQuery -Servers @($script:SelectedServer) -IsRescan $true
})

#endregion

#region -- Clear button ------------------------------------------------------

$btnClearTiles.Add_Click({
    $script:PollTimer.Stop()
    if ($script:RebootTimer) { $script:RebootTimer.Stop(); $script:RebootTimer.Dispose(); $script:RebootTimer = $null }
    foreach ($j in $script:ActiveJobs.Values) { Remove-Job -Job $j -Force -ErrorAction SilentlyContinue }
    $script:ActiveJobs     = @{}
    $script:PendingCount   = 0
    $pnlTiles.Controls.Clear()
    $script:ServerData     = @{}
    $script:SelectedServer = $null
    $btnRescan.Enabled     = $false
    $btnSetStatic.Enabled  = $false
    $btnSetStatic.BackColor = $clrBtnDisabled
    $btnSetStatic.ForeColor = $clrMuted
    $btnSetDHCP.Enabled    = $false
    $btnSetDHCP.BackColor  = $clrBtnDisabled
    $btnSetDHCP.ForeColor  = $clrMuted
    $btnReboot.Enabled     = $false
    $rtbDetail.Clear()
    $lblDetail.Text        = 'IP Detail:  (click a tile above to view)'
    $statusLabel.Text      = 'Cleared.'
    $btnRun.Enabled        = $true
    $btnRun.Text           = '> Run'
})

#endregion

#region -- Set Static button + dialogue --------------------------------------

$btnSetStatic.Add_Click({
    if (-not $script:SelectedServer) { return }
    $srv  = $script:SelectedServer
    $data = $script:ServerData[$srv]

    # Build dialogue form - this is still on the UI thread (instant, no network)
    $dlg               = New-Object System.Windows.Forms.Form
    $dlg.Text          = "Set Static IP  -  $($srv.ToUpper())"
    $dlg.Size          = New-Object System.Drawing.Size(360, 310)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox   = $false
    $dlg.MinimizeBox   = $false
    $dlg.BackColor     = $clrFormBg
    $dlg.ForeColor     = $clrText
    $dlg.Font          = New-Object System.Drawing.Font('Consolas', 9)

    $fields = @(
        @{ Label = 'IP Address:';      Key = 'IP';   Default = $data.PrimaryIP   },
        @{ Label = 'Subnet Mask:';     Key = 'Sub';  Default = $data.PrimaryMask },
        @{ Label = 'Default Gateway:'; Key = 'DG';   Default = $data.PrimaryGW   },
        @{ Label = 'DNS 1:';           Key = 'DNS1'; Default = $data.PrimaryDNS1 },
        @{ Label = 'DNS 2:';           Key = 'DNS2'; Default = $data.PrimaryDNS2 }
    )

    $textBoxes = @{}
    $y = 14
    foreach ($f in $fields) {
        $lbl           = New-Object System.Windows.Forms.Label
        $lbl.Text      = $f.Label
        $lbl.Location  = New-Object System.Drawing.Point(12, ($y + 3))
        $lbl.Size      = New-Object System.Drawing.Size(130, 18)
        $lbl.ForeColor = $clrMuted
        $dlg.Controls.Add($lbl)

        $txt             = New-Object System.Windows.Forms.TextBox
        $txt.Text        = $f.Default
        $txt.Location    = New-Object System.Drawing.Point(150, $y)
        $txt.Size        = New-Object System.Drawing.Size(180, 22)
        $txt.BackColor   = $clrPanelBg
        $txt.ForeColor   = $clrGreen
        $txt.BorderStyle = 'FixedSingle'
        $dlg.Controls.Add($txt)
        $textBoxes[$f.Key] = $txt
        $y += 36
    }

    $btnApply               = New-Object System.Windows.Forms.Button
    $btnApply.Text          = 'Set IP'
    $btnApply.Location      = New-Object System.Drawing.Point(150, ($y + 8))
    $btnApply.Size          = New-Object System.Drawing.Size(180, 32)
    $btnApply.BackColor     = $clrGreen
    $btnApply.ForeColor     = $clrFormBg
    $btnApply.FlatStyle     = 'Flat'
    $btnApply.Font          = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
    $btnApply.DialogResult  = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnApply)
    $dlg.AcceptButton = $btnApply

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = 'Cancel'
    $btnCancel.Location     = New-Object System.Drawing.Point(12, ($y + 8))
    $btnCancel.Size         = New-Object System.Drawing.Size(130, 32)
    $btnCancel.BackColor    = $clrTileGrey
    $btnCancel.ForeColor    = $clrText
    $btnCancel.FlatStyle    = 'Flat'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    # Dialog is instant - no network involved, UI is fine here
    if ($dlg.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $newIP   = $textBoxes['IP'].Text.Trim()
    $newSub  = $textBoxes['Sub'].Text.Trim()
    $newDG   = $textBoxes['DG'].Text.Trim()
    $newDNS1 = $textBoxes['DNS1'].Text.Trim()
    $newDNS2 = $textBoxes['DNS2'].Text.Trim()

    if ([string]::IsNullOrWhiteSpace($newIP) -or [string]::IsNullOrWhiteSpace($newSub)) {
        [System.Windows.Forms.MessageBox]::Show('IP Address and Subnet Mask are required.','Validation Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # Convert mask to prefix - done here on UI thread before handing off
    $maskBytes  = $newSub.Split('.') | ForEach-Object { [Convert]::ToByte($_) }
    $maskBinary = ($maskBytes | ForEach-Object { [Convert]::ToString($_, 2).PadLeft(8,'0') }) -join ''
    $prefix     = ($maskBinary.ToCharArray() | Where-Object { $_ -eq '1' }).Count

    # Mark tile as working so user gets visual feedback immediately
    Update-TileState $srv 'pending' 'Setting Static...' ''
    Add-DetailLine $srv '' $clrMuted
    Add-DetailLine $srv "--- Set Static submitted $(Get-Date -Format 'HH:mm:ss') ---" $clrAmber
    Add-DetailLine $srv "  IP: $newIP / $newSub  GW: $newDG" $clrMuted
    $statusLabel.Text     = "Setting static IP on $srv (background)..."
    $btnSetStatic.Enabled = $false
    $btnSetDHCP.Enabled   = $false

    # Fire the network work off as a background job - UI stays responsive
    $actionSB = {
        param($Server, $IP, $Prefix, $GW, $DNS1, $DNS2)
        try {
            $sessionOpt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 25000 -CancelTimeout 5000
            Invoke-Command -ComputerName $Server -SessionOption $sessionOpt -ScriptBlock {
                param($IP, $Prefix, $GW, $DNS1, $DNS2)
                $cfg = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1
                if (-not $cfg) { $cfg = Get-NetIPConfiguration | Where-Object { $_.IPv4Address -ne $null } | Select-Object -First 1 }
                if (-not $cfg) { throw 'Could not find a suitable network adapter.' }
                $idx = $cfg.InterfaceIndex
                Remove-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -Confirm:$false -ErrorAction SilentlyContinue
                Remove-NetRoute     -InterfaceIndex $idx -AddressFamily IPv4 -Confirm:$false -ErrorAction SilentlyContinue
                New-NetIPAddress -InterfaceIndex $idx -IPAddress $IP -PrefixLength $Prefix -DefaultGateway $GW -ErrorAction Stop | Out-Null
                $dnsServers = @($DNS1) | Where-Object { $_ -ne '' }
                if ($DNS2 -ne '') { $dnsServers += $DNS2 }
                if ($dnsServers.Count -gt 0) {
                    Set-DnsClientServerAddress -InterfaceIndex $idx -ServerAddresses $dnsServers -ErrorAction SilentlyContinue
                }
            } -ArgumentList $IP, $Prefix, $GW, $DNS1, $DNS2 -ErrorAction Stop
            return [PSCustomObject]@{ Success = $true;  Error = '';   Server = $Server; Action = 'SetStatic'; IP = $IP; Sub = ''; GW = $GW; DNS1 = $DNS1; DNS2 = $DNS2 }
        } catch {
            return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message; Server = $Server; Action = 'SetStatic'; IP = $IP; Sub = ''; GW = $GW; DNS1 = $DNS1; DNS2 = $DNS2 }
        }
    }
    $job = Start-Job -ScriptBlock $actionSB -ArgumentList $srv, $newIP, $prefix, $newDG, $newDNS1, $newDNS2
    $script:ActionJobs[$srv] = @{ Job = $job; Action = 'SetStatic'; Sub = $newSub; StartTime = [datetime]::Now }
    $script:ActionPollTimer.Start()
})

#endregion

#region -- Set DHCP button ---------------------------------------------------

$btnSetDHCP.Add_Click({
    if (-not $script:SelectedServer) { return }
    $srv = $script:SelectedServer

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Switch $srv to DHCP?`n`nThis will remove the static IP configuration and enable DHCP on the primary adapter.`nThe server may briefly lose connectivity.",
        'Confirm Set DHCP',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    Update-TileState $srv 'pending' 'Setting DHCP...' ''
    Add-DetailLine $srv '' $clrMuted
    Add-DetailLine $srv "--- Set DHCP submitted $(Get-Date -Format 'HH:mm:ss') ---" $clrAmber
    $statusLabel.Text     = "Setting DHCP on $srv (background)..."
    $btnSetStatic.Enabled = $false
    $btnSetDHCP.Enabled   = $false

    $actionSB = {
        param($Server)
        try {
            $sessionOpt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 25000 -CancelTimeout 5000
            Invoke-Command -ComputerName $Server -SessionOption $sessionOpt -ScriptBlock {
                $cfg = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1
                if (-not $cfg) { $cfg = Get-NetIPConfiguration | Where-Object { $_.IPv4Address -ne $null } | Select-Object -First 1 }
                if (-not $cfg) { throw 'Could not find a suitable network adapter.' }
                $idx = $cfg.InterfaceIndex
                Set-NetIPInterface -InterfaceIndex $idx -AddressFamily IPv4 -Dhcp Enabled -ErrorAction Stop
                Remove-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -Confirm:$false -ErrorAction SilentlyContinue
                Remove-NetRoute     -InterfaceIndex $idx -AddressFamily IPv4 -Confirm:$false -ErrorAction SilentlyContinue
                Set-DnsClientServerAddress -InterfaceIndex $idx -ResetServerAddresses -ErrorAction SilentlyContinue
            } -ErrorAction Stop
            return [PSCustomObject]@{ Success = $true;  Error = ''; Server = $Server; Action = 'SetDHCP' }
        } catch {
            return [PSCustomObject]@{ Success = $false; Error = $_.Exception.Message; Server = $Server; Action = 'SetDHCP' }
        }
    }
    $job = Start-Job -ScriptBlock $actionSB -ArgumentList $srv
    $script:ActionJobs[$srv] = @{ Job = $job; Action = 'SetDHCP'; Sub = ''; StartTime = [datetime]::Now }
    $script:ActionPollTimer.Start()
})

#endregion

#region -- Reboot button -----------------------------------------------------

$btnReboot.Add_Click({
    if (-not $script:SelectedServer) { return }
    $srv = $script:SelectedServer

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Reboot $srv now ?`n`nThis will immediately restart the server.`nAuto-rescan will run 30 seconds after the command is sent.",
        'Confirm Reboot',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnRescan.Enabled    = $false
    $btnSetStatic.Enabled = $false
    $btnSetStatic.BackColor = $clrBtnDisabled
    $btnSetStatic.ForeColor = $clrMuted
    $btnSetDHCP.Enabled   = $false
    $btnSetDHCP.BackColor  = $clrBtnDisabled
    $btnSetDHCP.ForeColor  = $clrMuted
    $btnReboot.Enabled    = $false
    $statusLabel.Text     = "Rebooting $srv..."

    Add-DetailLine $srv '' $clrMuted
    Add-DetailLine $srv "--- Reboot initiated $(Get-Date -Format 'HH:mm:ss') ---" $clrAmber

    try {
        Invoke-Command -ComputerName $srv -ScriptBlock { Restart-Computer -Force } -ErrorAction Stop

        Add-DetailLine $srv '  Reboot command sent successfully.' $clrGreen
        Add-DetailLine $srv '  Auto-rescan will run in 30 seconds...' $clrMuted
        Update-TileState $srv 'pending' 'Rebooting...' 'Rescan in 30s'

        [System.Windows.Forms.MessageBox]::Show(
            "Reboot command sent to $srv .`n`nAuto-rescan will run in 30 seconds.",
            'Reboot Sent',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        Start-RebootRescanCountdown -ServerName $srv

    } catch {
        $errMsg = $_.Exception.Message
        Add-DetailLine $srv "  [ERROR] Reboot failed: $errMsg" $clrRed
        $statusLabel.Text = "Reboot failed on $srv"
        [System.Windows.Forms.MessageBox]::Show(
            "Reboot failed on $srv .`n`n$errMsg",
            'Reboot Failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $btnRescan.Enabled = $true
        $btnReboot.Enabled = $true
    }
})

#endregion

#region -- Load servers.conf -------------------------------------------------

# Resolve script directory using every available method, including current dir
$scriptRoot = $null
if     ($PSScriptRoot      -and $PSScriptRoot      -ne '') { $scriptRoot = $PSScriptRoot }
elseif ($PSCommandPath     -and $PSCommandPath      -ne '') { $scriptRoot = Split-Path -Parent $PSCommandPath }
elseif ($MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
# Always also check the current working directory as a last resort
$searchRoots = @()
if ($scriptRoot) { $searchRoots += $scriptRoot }
$searchRoots += (Get-Location).Path

foreach ($root in ($searchRoots | Select-Object -Unique)) {
    $confPath = Join-Path $root 'servers.conf'
    if (Test-Path $confPath) {
        $confServers = Get-Content $confPath |
                       ForEach-Object { $_.Trim() } |
                       Where-Object   { $_ -ne '' -and -not $_.StartsWith('#') } |
                       Select-Object  -Unique
        if ($confServers) {
            $txtServers.Text  = ($confServers -join "`r`n")
            $statusLabel.Text = "Loaded $($confServers.Count) server(s) from servers.conf  [$confPath]"
        }
        break
    }
}

#endregion

#region -- Clean up on close -------------------------------------------------

$form.Add_FormClosing({
    $script:PollTimer.Stop()
    $script:ActionPollTimer.Stop()
    if ($script:RebootTimer) { $script:RebootTimer.Stop(); $script:RebootTimer.Dispose() }
    foreach ($j in $script:ActiveJobs.Values) { Remove-Job -Job $j -Force -ErrorAction SilentlyContinue }
    foreach ($ji in $script:ActionJobs.Values) { Remove-Job -Job $ji.Job -Force -ErrorAction SilentlyContinue }
})

#endregion

[System.Windows.Forms.Application]::Run($form)
