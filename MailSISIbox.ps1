<#
MailSISIbox - move files between a local PC and an AVD session through the
text clipboard, for when drive redirection and file copy-paste are turned off.

  MailSISIbox.ps1                       open the window
  MailSISIbox.ps1 -Receive C:\Drop      console, wait for incoming files
  MailSISIbox.ps1 -Send C:\book.xlsx    console, send a file

The clipboard always works, so it is the control channel: over it the two sides
agree on the fastest transport they actually share - a direct TCP connection, a
direct copy over \\tsclient\, or the clipboard itself - and fall back down that
ladder automatically when one is closed off.

PROTOCOL.md has the wire format and the reasoning behind the tuning.
#>
param(
    [string[]]$Send,
    [string]$Receive,
    [switch]$UI,

    [int]$StartChunkKB = 1024,
    [int]$MaxChunkMB   = 32,
    [int]$MinChunkKB   = 256,

    [int]$StallTimeoutSec = 300,

    # First port the receiver listens on. Fixed, so a firewall rule only has to
    # be written once.
    [int]$TcpPort = 48731,

    # Force the slow path, for testing the fallback or when a transport misbehaves.
    [switch]$NoTcp,
    [switch]$NoTsclient,

    [Alias('Minimised')]
    [switch]$Minimized,

    [switch]$InstallStartup,
    [switch]$RemoveStartup,

    [switch]$Trace
)

$ErrorActionPreference = 'Stop'

#region start at login

function Get-StartupCmdPath {
    Join-Path ([Environment]::GetFolderPath('Startup')) 'MailSISIbox.cmd'
}

function Install-Startup {
    $cmd = Get-StartupCmdPath
    $dir = Split-Path $cmd
    if (-not (Test-Path $dir)) {
        Write-Host "Could not find your Startup folder ($dir)." -ForegroundColor Red
        return $false
    }
    $body = @"
@echo off
rem Starts MailSISIbox when you log in. Delete this file to stop that,
rem or run: MailSISIbox.ps1 -RemoveStartup
powershell -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "$PSCommandPath" -Minimized
"@
    Set-Content -Path $cmd -Value $body -Encoding ASCII
    Write-Host ''
    Write-Host 'MailSISIbox will now start automatically when you log in.' -ForegroundColor Green
    Write-Host "  it runs: $PSCommandPath" -ForegroundColor DarkGray
    Write-Host "  set up by: $cmd" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'To stop it starting automatically, run this with -RemoveStartup.' -ForegroundColor Gray
    Write-Host ''
    if ($PSCommandPath -notlike "$env:USERPROFILE*") {
        Write-Host 'Note: this script is not inside your user profile. On a shared AVD' -ForegroundColor Yellow
        Write-Host 'machine that means it may be gone next time you log in. Keeping it' -ForegroundColor Yellow
        Write-Host 'somewhere like Documents is safer.' -ForegroundColor Yellow
        Write-Host ''
    }
    return $true
}

function Remove-Startup {
    $cmd = Get-StartupCmdPath
    if (Test-Path $cmd) {
        Remove-Item $cmd -Force
        Write-Host ''
        Write-Host 'MailSISIbox will no longer start automatically.' -ForegroundColor Green
        Write-Host ''
    } else {
        Write-Host ''
        Write-Host 'It was not set to start automatically, so nothing changed.' -ForegroundColor Gray
        Write-Host ''
    }
    return $true
}

# Done before the STA relaunch below - neither needs a window or the clipboard.
if ($InstallStartup) { if (Install-Startup) { exit 0 } else { exit 1 } }
if ($RemoveStartup)  { if (Remove-Startup)  { exit 0 } else { exit 1 } }

#endregion

# WinForms clipboard only works on an STA thread.
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    # Base64 of UTF-16 rather than -File with arguments: file names with spaces,
    # commas or non-ASCII characters survive it intact.
    function Quote { param($s) "'" + ([string]$s -replace "'", "''") + "'" }

    $cmd = '& ' + (Quote $PSCommandPath)
    if ($Send)       { $cmd += ' -Send @(' + (($Send | ForEach-Object { Quote $_ }) -join ',') + ')' }
    if ($Receive)    { $cmd += ' -Receive ' + (Quote $Receive) }
    if ($Trace)      { $cmd += ' -Trace' }
    if ($Minimized)  { $cmd += ' -Minimized' }
    if ($NoTcp)      { $cmd += ' -NoTcp' }
    if ($NoTsclient) { $cmd += ' -NoTsclient' }
    $cmd += " -TcpPort $TcpPort -StartChunkKB $StartChunkKB -MaxChunkMB $MaxChunkMB"
    $cmd += " -MinChunkKB $MinChunkKB -StallTimeoutSec $StallTimeoutSec"

    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
    & powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -EncodedCommand $enc
    exit $LASTEXITCODE
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Polling the sequence number is nearly free; reading the text is not.
Add-Type -Namespace MailSISIbox -Name Native -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetClipboardSequenceNumber();
[DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")]
public static extern bool OpenClipboard(IntPtr hWndNewOwner);
[DllImport("user32.dll")]
public static extern bool CloseClipboard();
[DllImport("user32.dll")]
public static extern IntPtr GetClipboardData(uint uFormat);
[DllImport("kernel32.dll")]
public static extern UIntPtr GlobalSize(IntPtr hMem);
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(IntPtr hWnd);
'@

$script:PROTO    = 'SISI|1'
$script:PollMs   = 25
$script:MaxChunk = $MaxChunkMB * 1MB
$script:MinChunk = $MinChunkKB * 1KB

# Already-compressed containers - deflating these just burns CPU for nothing.
$script:SkipCompress = @('.xlsx','.xlsb','.docx','.pptx','.zip','.7z','.gz','.rar',
                         '.png','.jpg','.jpeg','.gif','.mp4','.mov','.pdf')

$script:LastSeq       = 0
$script:ClipReads     = 0
$script:ClipReadChars = 0
$script:Cancel        = $false

$script:TransportLabel  = ''
$script:TransportDetail = ''

$script:NoTcpTransport = [bool]$NoTcp
$script:NoTs           = [bool]$NoTsclient

#region output hooks

$script:OnStatus = {
    param($Text, $Level)
    $c = switch ($Level) { 'ok' {'Green'} 'warn' {'Yellow'} 'bad' {'Red'} 'dim' {'DarkGray'} default {'Gray'} }
    Write-Host $Text -ForegroundColor $c
}
$script:OnProgress = {
    param($Label, $Done, $Total, $Rate, $Chunk, $Eta)
    Write-Host ("`r  {0,5}%  {1} / {2}   {3}/s   chunk {4}   eta {5}   " -f `
        ([Math]::Round(($Done / [Math]::Max(1L,[int64]$Total)) * 100, 1)),
        (Format-Size $Done), (Format-Size $Total), (Format-Size $Rate),
        (Format-Size $Chunk), $Eta) -NoNewline
}
$script:UiPump = $null

function Set-Transport { param([string]$Label) $script:TransportLabel = $Label }
function Say         { param([string]$Text, [string]$Level = 'info') & $script:OnStatus $Text $Level }
function Show-Advance { param($Label, $Done, $Total, $Rate, $Chunk, $Eta) & $script:OnProgress $Label $Done $Total $Rate $Chunk $Eta }

#endregion

#region clipboard

# Put text on the clipboard. Retries, because the clipboard is one shared lock
# and RDP's sync agent grabs it constantly.
function Set-ClipText {
    param([string]$Text)
    for ($i = 0; $i -lt 6; $i++) {
        try {
            if ([string]::IsNullOrEmpty($Text)) { [Windows.Forms.Clipboard]::Clear() }
            else { [Windows.Forms.Clipboard]::SetText($Text) }
            $script:LastSeq = [MailSISIbox.Native]::GetClipboardSequenceNumber()
            return $true
        } catch { Start-Sleep -Milliseconds 60 }
    }
    return $false
}

# Size without reading the text in. -1 means "could not tell", which callers
# treat as "just read it".
function Get-ClipTextSize {
    for ($i = 0; $i -lt 3; $i++) {
        if ([MailSISIbox.Native]::OpenClipboard([IntPtr]::Zero)) {
            try {
                $h = [MailSISIbox.Native]::GetClipboardData(13)   # CF_UNICODETEXT
                if ($h -eq [IntPtr]::Zero) { return 0 }
                # A UIntPtr will not cast straight to a number.
                return [int64]([MailSISIbox.Native]::GlobalSize($h).ToUInt64())
            } catch { return -1 } finally { [void][MailSISIbox.Native]::CloseClipboard() }
        }
        Start-Sleep -Milliseconds 40
    }
    return -1
}

function Get-ClipText {
    for ($i = 0; $i -lt 6; $i++) {
        try {
            if ([Windows.Forms.Clipboard]::ContainsText()) {
                return [Windows.Forms.Clipboard]::GetText()
            }
            return ''
        } catch { Start-Sleep -Milliseconds 60 }
    }
    return ''
}

#endregion

#region protocol

# Frame: SISI|1|TYPE|SESSION|A|B|FLAGS|HASH|PAYLOAD
# A and B mean different things per type - see PROTOCOL.md.
function New-Frame {
    param([string]$Type, [string]$Session, $A = 0, $B = 0,
          [string]$Flags = '-', [string]$Hash = '-', [string]$Payload = '')
    return "$script:PROTO|$Type|$Session|$A|$B|$Flags|$Hash|$Payload"
}

function Read-Frame {
    param([string]$Text)
    if (-not $Text -or -not $Text.StartsWith('SISI|1|')) { return $null }
    $p = $Text.Split('|', 9)
    if ($p.Count -lt 9) { return $null }
    return [pscustomobject]@{
        Type = $p[2]; Session = $p[3]; A = $p[4]; B = $p[5]
        Flags = $p[6]; Hash = $p[7]; Payload = $p[8]
    }
}

function Wait-Frame {
    param([string[]]$Types, [string]$Session, [double]$TimeoutSec, [int64]$MaxBytes = 0)

    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        if ($script:Cancel) { return $null }
        $seq = [MailSISIbox.Native]::GetClipboardSequenceNumber()
        if ($seq -ne $script:LastSeq) {
            $script:LastSeq = $seq

            # Too big to be the frame we want, so don't pay to read it. Only
            # the idle wait sets this; a chunk mid-transfer is legitimately big.
            if ($MaxBytes -gt 0) {
                $size = Get-ClipTextSize
                if ($size -gt $MaxBytes) {
                    if ($script:UiPump) { & $script:UiPump }
                    Start-Sleep -Milliseconds $script:PollMs
                    continue
                }
            }

            $script:ClipReads++
            $txt = Get-ClipText
            $script:ClipReadChars += $txt.Length
            $f = Read-Frame $txt
            $txt = $null
            if ($f -and $Types -contains $f.Type) {
                if (-not $Session -or $f.Session -eq $Session) { return $f }
            }
        }
        if ($script:UiPump) { & $script:UiPump }
        Start-Sleep -Milliseconds $script:PollMs
    }
    return $null
}

# Re-checks for a reply before each re-assert, so one that just landed is never
# overwritten.
function Send-AndWait {
    param([string]$Frame, [string[]]$Expect, [string]$Session,
          [double]$TimeoutSec, [double]$RetryEverySec = 5)

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $first = $true
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        if ($script:Cancel) { return $null }
        if (-not $first) {
            $seq = [MailSISIbox.Native]::GetClipboardSequenceNumber()
            if ($seq -ne $script:LastSeq) {
                $script:LastSeq = $seq
                $f = Read-Frame (Get-ClipText)
                if ($f -and $Expect -contains $f.Type -and (-not $Session -or $f.Session -eq $Session)) {
                    return $f
                }
            }
        }
        $first = $false

        Set-ClipText $Frame | Out-Null
        $slice = [Math]::Min($RetryEverySec, $TimeoutSec - $sw.Elapsed.TotalSeconds)
        if ($slice -le 0) { break }
        $f = Wait-Frame -Types $Expect -Session $Session -TimeoutSec $slice
        if ($f) { return $f }
    }
    return $null
}

# Hold the closing frame until the sender stops asking, in case one got clobbered.
function Publish-Final {
    param([string]$Frame, [string]$Session, [double]$LingerSec = 12)
    Set-ClipText $Frame | Out-Null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $quiet = 0
    while ($sw.Elapsed.TotalSeconds -lt $LingerSec) {
        if (Wait-Frame -Types @('DONE') -Session $Session -TimeoutSec 2) {
            Set-ClipText $Frame | Out-Null
            $quiet = 0
        } else {
            $quiet++
            if ($quiet -ge 2) { break }
        }
    }
}

#endregion

#region chunk helpers

# Truncated on purpose - the whole-file hash at the end is the real check.
function Get-ChunkHash {
    param([byte[]]$Bytes, [int]$Length)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $h = $sha.ComputeHash($Bytes, 0, $Length)
        return [BitConverter]::ToString($h, 0, 8).Replace('-', '')
    } finally { $sha.Dispose() }
}

function Get-FileHashHex {
    param([string]$Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    $fs = [IO.File]::OpenRead($Path)
    try {
        return [BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '')
    } finally { $fs.Dispose(); $sha.Dispose() }
}

function Compress-Bytes {
    param([byte[]]$Bytes, [int]$Length)
    $ms = New-Object IO.MemoryStream
    $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Compress, $true)
    try { $ds.Write($Bytes, 0, $Length) } finally { $ds.Dispose() }
    $out = $ms.ToArray()
    $ms.Dispose()
    # The comma matters: without it PowerShell unrolls the array into the
    # pipeline and the caller gets megabytes of boxed bytes instead.
    return ,$out
}

function Expand-Bytes {
    param([byte[]]$Bytes, [int]$ExpectedLength)
    $ms = New-Object IO.MemoryStream(,$Bytes)
    $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
    $out = New-Object byte[] $ExpectedLength
    try {
        $got = 0
        while ($got -lt $ExpectedLength) {
            $n = $ds.Read($out, $got, $ExpectedLength - $got)
            if ($n -le 0) { break }
            $got += $n
        }
    } finally { $ds.Dispose(); $ms.Dispose() }
    return ,$out
}

# FileStream.Read is allowed to return short, so loop until the buffer is full.
function Read-Exact {
    param([IO.FileStream]$Stream, [byte[]]$Buffer, [int]$Count)
    $got = 0
    while ($got -lt $Count) {
        $n = $Stream.Read($Buffer, $got, $Count - $got)
        if ($n -le 0) { break }
        $got += $n
    }
    return $got
}

# Base64 of a big chunk lands on the large object heap, reclaimed only by a full
# collection - left alone the process runs past 4 GB.
function Clear-LargeAllocations {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Format-Size {
    param([double]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

#endregion

#region transport ladder

# Loopback is included on purpose: it is what lets both ends run on one machine.
function Get-LocalAddresses {
    $out = @()
    try {
        foreach ($ip in [Net.Dns]::GetHostAddresses([Net.Dns]::GetHostName())) {
            if ($ip.AddressFamily -eq 'InterNetwork' -and -not $ip.ToString().StartsWith('169.254')) {
                $out += $ip.ToString()
            }
        }
    } catch {}
    $out += '127.0.0.1'
    return @($out | Select-Object -Unique)
}

function Test-Tsclient {
    if ($script:NoTs) { return $false }
    try { return [bool](Get-ChildItem '\\tsclient\' -ErrorAction SilentlyContinue) } catch { return $false }
}

# Advert lines the receiver puts in its READY frame.
#   tcp=<ip>,<ip>:<port>   receiver is listening here
#   ts=1                   receiver can see \\tsclient\
#   dir=<abs path>         receiver's destination folder
function New-Advert {
    param($Listener, [string]$Folder)
    $lines = @()
    if ($Listener) {
        $port = $Listener.LocalEndpoint.Port
        $ips = @(Get-LocalAddresses)
        if ($ips.Count) { $lines += ('tcp=' + ($ips -join ',') + ':' + $port) }
    }
    if (Test-Tsclient) { $lines += 'ts=1' }
    $lines += "dir=$Folder"
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($lines -join "`n"))
}

function Read-Advert {
    param([string]$B64)
    $a = [pscustomobject]@{ TcpHosts = @(); TcpPort = 0; Tsclient = $false; Dir = '' }
    if (-not $B64 -or $B64 -eq '-') { return $a }
    try { $txt = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($B64)) } catch { return $a }
    foreach ($line in $txt.Split("`n")) {
        $i = $line.IndexOf('=')
        if ($i -lt 1) { continue }
        $k = $line.Substring(0, $i); $v = $line.Substring($i + 1)
        switch ($k) {
            'tcp' {
                $c = $v.LastIndexOf(':')
                if ($c -gt 0) {
                    $a.TcpHosts = $v.Substring(0, $c).Split(',')
                    $a.TcpPort  = [int]$v.Substring($c + 1)
                }
            }
            'ts'  { $a.Tsclient = ($v -eq '1') }
            'dir' { $a.Dir = $v }
        }
    }
    return $a
}

# A fixed small range, so a firewall rule only has to be written once.
function Start-DataListener {
    param([int]$BasePort, [int]$Range = 8)
    if ($script:NoTcpTransport) { return $null }
    for ($p = $BasePort; $p -lt ($BasePort + $Range); $p++) {
        try {
            $l = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Any, $p)
            $l.Start()
            return $l
        } catch { }
    }
    return $null
}

function Stop-DataListener {
    param($Listener)
    if ($Listener) { try { $Listener.Stop() } catch {} }
}

function Connect-WithTimeout {
    param([string]$Address, [int]$Port, [int]$TimeoutMs = 1200)
    $c = New-Object Net.Sockets.TcpClient
    try {
        $iar = $c.BeginConnect($Address, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { $c.Close(); return $null }
        $c.EndConnect($iar)
        $c.NoDelay = $true
        $c.SendBufferSize = 1MB
        $c.ReceiveBufferSize = 1MB
        return $c
    } catch { try { $c.Close() } catch {}; return $null }
}

# C:\Drop on the other machine is \\tsclient\C\Drop from inside the session.
# Drive letters only - a UNC or network path will not be there under one.
function Convert-ToTsclientPath {
    param([string]$LocalPath)
    if (-not $LocalPath -or $LocalPath.Length -lt 3) { return $null }
    if ($LocalPath[1] -ne ':' -or $LocalPath[2] -ne '\') { return $null }
    return '\\tsclient\' + $LocalPath[0] + $LocalPath.Substring(2)
}

# Carries the content hash, so the file's own length is a trustworthy resume point.
function Get-DropName {
    param([string]$Name, [string]$FileHash)
    return "$Name.$($FileHash.Substring(0,8)).sisidrop"
}

# The byte pump for both fast transports, reporting the same way the clipboard
# loop does so callers need not care which is live.
function Copy-Bytes {
    param([IO.Stream]$In, [IO.Stream]$Out, [int64]$Count, [string]$Name,
          [int64]$Base, [int64]$Total, [string]$PartPath = '', [string]$FileHash = '',
          [int]$BufferKB = 4096, [scriptblock]$Heartbeat = $null)

    $buf = New-Object byte[] ($BufferKB * 1KB)
    $moved = 0L
    $lastPart = 0L
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $lastTick = 0.0
    $lastBeat = 0.0

    while ($moved -lt $Count) {
        if ($script:Cancel) { break }
        $want = [int][Math]::Min([int64]$buf.Length, $Count - $moved)
        $n = $In.Read($buf, 0, $want)
        if ($n -le 0) { break }
        $Out.Write($buf, 0, $n)
        $moved += $n

        if ($PartPath -and ($moved - $lastPart) -ge 32MB) {
            $Out.Flush()
            Set-Content -Path $PartPath -Value "$FileHash`n$($Base + $moved)`n$Total" -NoNewline
            $lastPart = $moved
        }

        $el = $sw.Elapsed.TotalSeconds
        if (($el - $lastTick) -ge 0.2) {
            $lastTick = $el
            $rate = $moved / [Math]::Max(0.001, $el)
            $eta = if ($rate -gt 0) { [TimeSpan]::FromSeconds(($Count - $moved) / $rate).ToString('hh\:mm\:ss') } else { '--' }
            Show-Advance $Name ($Base + $moved) $Total $rate $buf.Length $eta
            if ($script:UiPump) { & $script:UiPump }

            if ($Heartbeat -and ($el - $lastBeat) -ge 5) {
                $lastBeat = $el
                & $Heartbeat ($Base + $moved)
            }
        }
    }

    $Out.Flush()
    $sw.Stop()
    return $moved
}

#endregion

#region transport: direct TCP

# Plain ASCII handshake, so neither end needs a byte count up front. Everything
# after it is raw file bytes.
function Write-Line {
    param($Stream, [string]$Text)
    $b = [Text.Encoding]::ASCII.GetBytes($Text + "`n")
    $Stream.Write($b, 0, $b.Length)
    $Stream.Flush()
}

function Read-Line {
    param($Stream, [int]$TimeoutMs = 30000)
    $old = $Stream.ReadTimeout
    $Stream.ReadTimeout = $TimeoutMs
    $sb = New-Object Text.StringBuilder
    try {
        while ($true) {
            $b = $Stream.ReadByte()
            if ($b -lt 0) { return $null }
            if ($b -eq 10) { break }
            [void]$sb.Append([char]$b)
            if ($sb.Length -gt 512) { return $null }
        }
    } catch { return $null } finally { try { $Stream.ReadTimeout = $old } catch {} }
    return $sb.ToString()
}

function Send-ViaTcp {
    param([string]$Path, [string]$Session, [int64]$Size, [string]$Name,
          $Advert, [ref]$Offset)

    $client = $null
    foreach ($h in $Advert.TcpHosts) {
        $client = Connect-WithTimeout -Address $h -Port $Advert.TcpPort
        if ($client) { $script:TransportDetail = ('{0}:{1}' -f $h, $Advert.TcpPort); break }
    }
    if (-not $client) { return $false }

    try {
        $ns = $client.GetStream()
        Write-Line $ns "SISI|1|TCP|$Session|$($Offset.Value)"

        $reply = Read-Line $ns
        if (-not $reply -or -not $reply.StartsWith('OK|')) { return $false }
        $at = [int64]$reply.Split('|')[1]
        $Offset.Value = $at

        $fs = [IO.File]::OpenRead($Path)
        try {
            $fs.Seek($at, 'Begin') | Out-Null
            $moved = Copy-Bytes -In $fs -Out $ns -Count ($Size - $at) -Name $Name -Base $at -Total $Size
            if ($moved -ne ($Size - $at)) { return $false }
        } finally { $fs.Dispose() }

        # Wait for the receiver to finish writing, so DONE never goes out over a
        # half-written file.
        $end = Read-Line $ns -TimeoutMs 120000
        if ($end -ne 'GOT') { return $false }
        $Offset.Value = $Size
        return $true
    } catch {
        Say "  direct connection failed: $($_.Exception.Message)" 'warn'
        return $false
    } finally {
        try { $client.Close() } catch {}
    }
}

function Receive-ViaTcp {
    param($Listener, [string]$Session, [IO.FileStream]$Out, [int64]$Size,
          [string]$Name, [string]$PartPath, [string]$FileHash, [ref]$Offset,
          [int]$WaitSec = 30)

    if (-not $Listener) { return $false }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (-not $Listener.Pending()) {
        if ($script:Cancel -or $sw.Elapsed.TotalSeconds -gt $WaitSec) { return $false }
        if ($script:UiPump) { & $script:UiPump }
        Start-Sleep -Milliseconds 50
    }

    $client = $null
    try {
        $client = $Listener.AcceptTcpClient()
        $client.NoDelay = $true
        $client.ReceiveBufferSize = 1MB
        $ns = $client.GetStream()

        # The session id only exists on the shared clipboard, so it is what
        # separates the real sender from anything else reaching this port.
        $hello = Read-Line $ns -TimeoutMs 10000
        if (-not $hello) { return $false }
        $p = $hello.Split('|')
        if ($p.Count -lt 4 -or $p[2] -ne 'TCP' -or $p[3] -ne $Session) {
            Say '  rejected a connection with the wrong session id' 'warn'
            return $false
        }

        $at = $Offset.Value
        $Out.Seek($at, 'Begin') | Out-Null
        Write-Line $ns "OK|$at"

        $moved = Copy-Bytes -In $ns -Out $Out -Count ($Size - $at) -Name $Name `
                            -Base $at -Total $Size -PartPath $PartPath -FileHash $FileHash
        if ($moved -ne ($Size - $at)) { return $false }

        $Offset.Value = $Size
        Set-Content -Path $PartPath -Value "$FileHash`n$Size`n$Size" -NoNewline
        Write-Line $ns 'GOT'
        return $true
    } catch {
        Say "  direct connection failed: $($_.Exception.Message)" 'warn'
        return $false
    } finally {
        try { $client.Close() } catch {}
    }
}

#endregion

#region transport: \\tsclient\ drive redirection

function Send-ViaTsclient {
    param([string]$Path, [string]$Name, [int64]$Size, [string]$FileHash, [string]$RemoteDir)

    $dir = Convert-ToTsclientPath $RemoteDir
    if (-not $dir) { return $false }
    if (-not (Test-Path -LiteralPath $dir)) { return $false }

    $drop = Join-Path $dir (Get-DropName $Name $FileHash)
    $at = 0L
    if (Test-Path -LiteralPath $drop) {
        $at = (Get-Item -LiteralPath $drop).Length
        if ($at -gt $Size) { $at = 0L }
        if ($at -gt 0) { Say "  resuming the direct copy at $(Format-Size $at)" 'warn' }
    }
    if ($at -eq $Size) { return $true }

    $in = [IO.File]::OpenRead($Path)
    $out = [IO.File]::Open($drop, 'OpenOrCreate', 'Write')
    try {
        $in.Seek($at, 'Begin') | Out-Null
        $out.Seek($at, 'Begin') | Out-Null
        $moved = Copy-Bytes -In $in -Out $out -Count ($Size - $at) -Name $Name -Base $at -Total $Size
        return ($moved -eq ($Size - $at))
    } catch {
        Say "  direct copy failed: $($_.Exception.Message)" 'warn'
        return $false
    } finally { $in.Dispose(); $out.Dispose() }
}

function Receive-ViaTsclient {
    param([string]$RemotePath, [IO.FileStream]$Out, [int64]$Size, [string]$Name,
          [string]$PartPath, [string]$FileHash, [ref]$Offset, [string]$Session)

    $src = Convert-ToTsclientPath $RemotePath
    if (-not $src) { return $false }
    if (-not (Test-Path -LiteralPath $src)) { return $false }

    $in = [IO.File]::OpenRead($src)
    try {
        $at = $Offset.Value
        $in.Seek($at, 'Begin') | Out-Null
        $Out.Seek($at, 'Begin') | Out-Null
        # Safe as a closure, unlike the window's handlers: it captures $Session,
        # touches no $script: state, and runs on this thread, not the message loop.
        $beat = { param($n) Set-ClipText (New-Frame -Type 'BUSY' -Session $Session -A $n) | Out-Null }.GetNewClosure()
        $moved = Copy-Bytes -In $in -Out $Out -Count ($Size - $at) -Name $Name `
                            -Base $at -Total $Size -PartPath $PartPath -FileHash $FileHash `
                            -Heartbeat $beat
        if ($moved -ne ($Size - $at)) { return $false }
        $Offset.Value = $Size
        Set-Content -Path $PartPath -Value "$FileHash`n$Size`n$Size" -NoNewline
        return $true
    } catch {
        Say "  direct read failed: $($_.Exception.Message)" 'warn'
        return $false
    } finally { $in.Dispose() }
}

#endregion

#region send

# One chunk out, one ack back. A round trip costs the same whatever it carries,
# so chunk size and compression are tuned against the live link as it runs.
function Send-ViaClipboard {
    param([string]$Path, [string]$Session, [int64]$Size, [string]$Name,
          [string]$Extension, [ref]$Offset)

    $at  = $Offset.Value
    $chunk   = [int]($StartChunkKB * 1KB)
    $rateFor = @{}          # chunk size -> best MB/s seen at that size
    $chunkNo = 0
    $sampled = $false

    # Compression is settled first, then chunk size - one at a time, or neither
    # measurement means anything.
    $canCompress   = $script:SkipCompress -notcontains $Extension.ToLower()
    $compressMode  = if ($canCompress) { 'probe' } else { 'off' }
    $compressStats = @{ on = @{ bytes = 0L; sec = 0.0; n = 0 }; off = @{ bytes = 0L; sec = 0.0; n = 0 } }
    $useCompress   = $false
    $probing       = ($compressMode -eq 'off')   # chunk probing waits its turn

    $buf   = New-Object byte[] $script:MaxChunk
    $fs    = [IO.File]::OpenRead($Path)
    $start = [Diagnostics.Stopwatch]::StartNew()
    $sentBytes = 0L

    try {
        $fs.Seek($at, 'Begin') | Out-Null
        while ($at -lt $Size) {
            if ($script:Cancel) { Say '  cancelled' 'warn'; return $false }

            $chunkSw = [Diagnostics.Stopwatch]::StartNew()
            $want = [Math]::Min([int64]$chunk, $Size - $at)
            $got  = Read-Exact -Stream $fs -Buffer $buf -Count ([int]$want)
            if ($got -le 0) { throw 'unexpected end of file' }

            $script:ClipReads = 0; $script:ClipReadChars = 0
            $tRead = $chunkSw.Elapsed.TotalMilliseconds

            # One small sample decides it: if that barely shrinks, nothing here will.
            if ($canCompress -and -not $sampled) {
                $sampled = $true
                $sampleLen = [int][Math]::Min(256KB, $got)
                $sample = Compress-Bytes -Bytes $buf -Length $sampleLen
                if ($sample.Length -gt ($sampleLen * 0.9)) {
                    $canCompress = $false
                    $compressMode = 'off'
                    $probing = $true
                }
                $sample = $null
            }

            if ($compressMode -eq 'probe') { $useCompress = (($chunkNo % 2) -eq 0) }
            else { $useCompress = ($compressMode -eq 'on') }

            $hash  = Get-ChunkHash -Bytes $buf -Length $got
            $tHash = $chunkSw.Elapsed.TotalMilliseconds
            $flags = '-'
            $body  = $buf
            $bodyLen = $got

            if ($useCompress) {
                $z = Compress-Bytes -Bytes $buf -Length $got
                # Only worth it if it actually saved something meaningful.
                if ($z.Length -lt ($got * 0.9)) { $flags = 'z'; $body = $z; $bodyLen = $z.Length }
            }

            $tComp = $chunkSw.Elapsed.TotalMilliseconds
            $payload = [Convert]::ToBase64String($body, 0, $bodyLen)
            $frame = New-Frame -Type 'DATA' -Session $Session -A $at -B $got -Flags $flags -Hash $hash -Payload $payload
            $tFrame = $chunkSw.Elapsed.TotalMilliseconds

            # Timeouts have to scale with chunk size or big chunks look like failures.
            $timeout = 20 + ($got / 1MB) * 3
            $retry   = [Math]::Max(8, ($got / 1MB) * 2.5)
            $reply = Send-AndWait -Frame $frame -Expect @('ACK','NAK') -Session $Session `
                                  -TimeoutSec $timeout -RetryEverySec $retry

            if ($Trace) {
                Write-Host ("`n    [{0,5:N1}MB] read {1,5:N0} hash {2,5:N0} comp {3,5:N0} b64 {4,5:N0} wait {5,6:N0} ms | reads {6}" -f `
                    ($got/1MB), $tRead, ($tHash-$tRead), ($tComp-$tHash), ($tFrame-$tComp),
                    ($chunkSw.Elapsed.TotalMilliseconds-$tFrame), $script:ClipReads) -ForegroundColor DarkCyan
            }

            $payload = $null; $frame = $null; $body = $null; $z = $null
            $chunkNo++
            if (($chunkNo % 8) -eq 0) { Clear-LargeAllocations }

            if (-not $reply -or $reply.Type -eq 'NAK') {
                # Back off and retry from wherever the receiver says it actually is.
                $chunk = [Math]::Max($script:MinChunk, [int]($chunk / 2))
                $probing = $false
                if ($reply -and $reply.Type -eq 'NAK') { $at = [int64]$reply.A }
                $Offset.Value = $at
                $fs.Seek($at, 'Begin') | Out-Null
                Say "  retrying at $(Format-Size $at), chunk now $(Format-Size $chunk)" 'warn'
                continue
            }

            $at = [int64]$reply.A
            $Offset.Value = $at
            $fs.Seek($at, 'Begin') | Out-Null
            $sentBytes += $got

            $chunkSw.Stop()
            $secs = [Math]::Max(0.001, $chunkSw.Elapsed.TotalSeconds)
            $thisRate = $got / $secs

            # Whichever setting moves more original bytes per second wins: 'off'
            # on a fast link, where the CPU dominates, 'on' over real RDP.
            if ($compressMode -eq 'probe') {
                $k = if ($useCompress) { 'on' } else { 'off' }
                $compressStats[$k].bytes += $got
                $compressStats[$k].sec   += $secs
                $compressStats[$k].n++
                if ($compressStats.on.n -ge 2 -and $compressStats.off.n -ge 2) {
                    $rOn  = $compressStats.on.bytes  / $compressStats.on.sec
                    $rOff = $compressStats.off.bytes / $compressStats.off.sec
                    $compressMode = if ($rOn -gt $rOff) { 'on' } else { 'off' }
                    $probing = $true
                }
            }

            # Only trust chunk-size numbers once compression has stopped changing.
            if ($compressMode -ne 'probe') {
                if (-not $rateFor.ContainsKey($chunk) -or $thisRate -gt $rateFor[$chunk]) {
                    $rateFor[$chunk] = $thisRate
                }
            }

            # Grow only while growing actually helps, rather than assuming which
            # kind of link this is.
            if ($probing -and $chunk -lt $script:MaxChunk) {
                $prev = [int]($chunk / 2)
                if (-not $rateFor.ContainsKey($prev) -or $rateFor[$chunk] -ge ($rateFor[$prev] * 1.05)) {
                    $chunk = [Math]::Min($script:MaxChunk, $chunk * 2)
                } else {
                    # The step up made things worse - go back and settle there.
                    $chunk = $prev
                    $probing = $false
                }
            } elseif ($probing) {
                $probing = $false
            }

            $rate = $sentBytes / [Math]::Max(0.001, $start.Elapsed.TotalSeconds)
            $eta  = if ($rate -gt 0) { [TimeSpan]::FromSeconds(($Size - $at) / $rate).ToString('hh\:mm\:ss') } else { '--' }
            Show-Advance $Name $at $Size $rate $chunk $eta
        }
        return $true
    } finally {
        $fs.Dispose()
    }
}

# Which transports are worth trying, best first. Candidates to attempt in order
# rather than a verdict, because the only way to find out is to try.
function Get-TransportPlan {
    param($Advert, [string]$SourcePath)

    $plan = @()

    if ($Advert.TcpHosts.Count -gt 0) {
        $plan += [pscustomobject]@{ Kind = 'tcp'; Detail = ''; Label = 'direct connection' }
    }

    # Whichever side is inside the session can see the other's disk: that is a
    # push from in there, a pull from out here.
    if ((Test-Tsclient) -and (Convert-ToTsclientPath $Advert.Dir)) {
        $plan += [pscustomobject]@{ Kind = 'tspush'; Detail = 'push'; Label = 'direct drive copy' }
    } elseif ($Advert.Tsclient -and (Convert-ToTsclientPath $SourcePath)) {
        $plan += [pscustomobject]@{ Kind = 'tspull'; Detail = $SourcePath; Label = 'direct drive copy' }
    }

    $plan += [pscustomobject]@{ Kind = 'clip'; Detail = ''; Label = 'clipboard' }
    return $plan
}

# On a pull the receiver does all the work, so this just watches its BUSY frames
# to tell a long transfer from a dead one.
function Wait-ForPull {
    param([string]$Session, [int64]$Size, [string]$Name, [ref]$Offset)

    $quiet = [Diagnostics.Stopwatch]::StartNew()
    while ($quiet.Elapsed.TotalSeconds -lt $StallTimeoutSec) {
        if ($script:Cancel) { return $false }

        $f = Wait-Frame -Types @('ACK','READY','BUSY') -Session $Session -TimeoutSec 5
        if (-not $f) { continue }

        $at = [int64]$f.A
        $Offset.Value = $at

        if ($f.Type -eq 'BUSY') {
            $quiet.Restart()
            Show-Advance $Name $at $Size 0 0 '--'
            continue
        }
        # An ACK for the whole file means it landed; a READY means it gave up
        # and is advertising for another transport.
        return ($f.Type -eq 'ACK' -and $at -ge $Size)
    }
    Say '  the other side stopped reporting progress' 'warn'
    return $false
}

# Send one file: offer it, agree a route, try each route in turn, and wait to
# hear the far side verify what landed.
function Send-OneFile {
    param([string]$Path)

    $file = Get-Item -LiteralPath $Path
    $size = $file.Length
    $name = $file.Name

    Say ''
    Say "Sending $name  ($(Format-Size $size))" 'info'
    $fileHash = Get-FileHashHex $file.FullName

    $session = [guid]::NewGuid().ToString('N').Substring(0, 8)

    # Offer, and keep re-offering until the far side answers - it may not be up yet.
    $meta = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$name`n$size`n$fileHash"))
    $offer = New-Frame -Type 'OFFER' -Session $session -A $size -B 0 -Payload $meta

    Say '  waiting for the other side...' 'dim'
    $ready = Send-AndWait -Frame $offer -Expect @('READY') -Session $session `
                          -TimeoutSec $StallTimeoutSec -RetryEverySec 3
    if (-not $ready) { Say '  timed out - is MailSISIbox running over there?' 'bad'; return $false }

    $offset = [int64]$ready.A
    if ($offset -gt 0) { Say "  resuming at $(Format-Size $offset)" 'warn' }

    $advert = Read-Advert $ready.Payload
    $plan   = Get-TransportPlan -Advert $advert -SourcePath $file.FullName
    $start  = [Diagnostics.Stopwatch]::StartNew()
    $moved  = $false

    foreach ($step in $plan) {
        if ($script:Cancel) { break }
        if ($offset -ge $size) { $moved = $true; break }

        # Tell the receiver which channel to listen on. Its reply is the
        # authoritative offset - a failed attempt may have moved it.
        $script:TransportDetail = ''
        $detail = if ($step.Detail) { [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($step.Detail)) } else { '' }
        $use = New-Frame -Type 'USE' -Session $session -A $offset -Flags $step.Kind -Payload $detail
        $set = Send-AndWait -Frame $use -Expect @('SET') -Session $session -TimeoutSec 60 -RetryEverySec 1.5
        if (-not $set) { Say '  the other side did not answer - stopping' 'bad'; return $false }
        if ($set.Flags -ne 'ok') { continue }          # receiver cannot use this one

        $offset = [int64]$set.A
        $ref = [ref]$offset

        Set-Transport $step.Label
        Say "  route: $($step.Label)" 'dim'

        switch ($step.Kind) {
            'tcp'    { $moved = Send-ViaTcp -Path $file.FullName -Session $session -Size $size -Name $name -Advert $advert -Offset $ref }
            'tspush' { $moved = Send-ViaTsclient -Path $file.FullName -Name $name -Size $size -FileHash $fileHash -RemoteDir $advert.Dir
                       if ($moved) { $offset = $size; $ref.Value = $size } }
            # The receiver reads it off our disk itself; we just wait to hear how
            # that went.
            'tspull' { $moved = Wait-ForPull -Session $session -Size $size -Name $name -Offset $ref }
            'clip'   { $moved = Send-ViaClipboard -Path $file.FullName -Session $session -Size $size -Name $name -Extension $file.Extension -Offset $ref }
        }
        $offset = $ref.Value

        if ($moved) {
            if ($script:TransportDetail) { Say "  via $($script:TransportDetail)" 'dim' }
            break
        }
        if ($script:Cancel) { break }
        Say "  $($step.Label) did not work here - trying the next option" 'warn'
    }

    if ($script:Cancel) { Say '  cancelled' 'warn'; return $false }
    if (-not $moved) { Say '  no transport could move the file' 'bad'; return $false }

    # Whatever carried the bytes, it is only finished once the receiver has
    # hashed it and said so.
    $fin = Send-AndWait -Frame (New-Frame -Type 'DONE' -Session $session -A $size) `
                        -Expect @('FIN') -Session $session -TimeoutSec 3600 -RetryEverySec 3

    if ($fin -and $fin.Flags -eq 'ok') {
        Say ''
        Say ("  sent {0} in {1} ({2}/s)" -f (Format-Size $size),
            $start.Elapsed.ToString('hh\:mm\:ss'),
            (Format-Size ($size / [Math]::Max(0.001,$start.Elapsed.TotalSeconds)))) 'ok'
        return $true
    }
    Say '  FAILED - hash mismatch on the far side' 'bad'
    return $false
}

#endregion

#region receive

# The idle wait: sit until somebody offers us a file.
function Wait-ForOffer {
    param([double]$TimeoutSec)
    return Wait-Frame -Types @('OFFER') -Session $null -TimeoutSec $TimeoutSec -MaxBytes 64KB
}

# Takes DATA frames until the file is whole or the sender says it is done.
function Receive-ViaClipboard {
    param([IO.FileStream]$Out, [string]$Session, [int64]$Size, [string]$Name,
          [string]$PartPath, [string]$FileHash, [ref]$Offset, [ref]$SawDone)

    $at = $Offset.Value
    $start = [Diagnostics.Stopwatch]::StartNew()
    $gotBytes = 0L
    $chunkNo = 0

    # Nothing to publish on the first pass: the sender's first chunk is already
    # on its way, and writing now would wipe it.
    $outFrame = $null

    while ($at -lt $Size) {
        if ($script:Cancel) { Say '  cancelled' 'warn'; return $false }

        if ($outFrame) {
            $f = Send-AndWait -Frame $outFrame -Expect @('DATA','DONE') -Session $Session `
                              -TimeoutSec $StallTimeoutSec -RetryEverySec 10
        } else {
            $f = Wait-Frame -Types @('DATA','DONE') -Session $Session -TimeoutSec $StallTimeoutSec
        }
        if (-not $f) { Say '  stalled - giving up' 'bad'; return $false }

        if ($f.Type -eq 'DONE') { $SawDone.Value = $true; break }

        # Where the sender says this chunk goes, against where we actually are.
        $frameAt = [int64]$f.A
        $len     = [int]$f.B

        # Duplicate of something we already wrote - re-ack so the sender moves on.
        if ($frameAt -lt $at) {
            $outFrame = New-Frame -Type 'ACK' -Session $Session -A $at
            continue
        }
        # A gap means we lost a frame; tell the sender where we actually are.
        if ($frameAt -gt $at) {
            $outFrame = New-Frame -Type 'NAK' -Session $Session -A $at
            continue
        }

        $raw = [Convert]::FromBase64String($f.Payload)
        if ($f.Flags -eq 'z') { $raw = Expand-Bytes -Bytes $raw -ExpectedLength $len }

        if ((Get-ChunkHash -Bytes $raw -Length $len) -ne $f.Hash) {
            $outFrame = New-Frame -Type 'NAK' -Session $Session -A $at
            continue
        }

        $Out.Write($raw, 0, $len)
        $Out.Flush()
        $at += $len
        $gotBytes += $len
        $Offset.Value = $at

        # The sidecar is what makes resume work after a crash or disconnect.
        Set-Content -Path $PartPath -Value "$FileHash`n$at`n$Size" -NoNewline

        $outFrame = New-Frame -Type 'ACK' -Session $Session -A $at

        $rate = $gotBytes / [Math]::Max(0.001, $start.Elapsed.TotalSeconds)
        $eta  = if ($rate -gt 0) { [TimeSpan]::FromSeconds(($Size - $at) / $rate).ToString('hh\:mm\:ss') } else { '--' }
        Show-Advance $Name $at $Size $rate $len $eta

        $raw = $null; $f = $null
        $chunkNo++
        if (($chunkNo % 8) -eq 0) { Clear-LargeAllocations }
    }

    # Leave the last ack on the clipboard so the sender can stop re-sending.
    if ($outFrame) { Set-ClipText $outFrame | Out-Null }
    return $true
}

# A push makes no clipboard traffic at all while it runs, so watch the file grow
# instead and only give up when it stops growing.
function Wait-ForPush {
    param([string]$Path, [int64]$Size, [string]$Name, [string]$Session, [ref]$SawDone)

    $lastLen  = -1L
    $lastMove = [Diagnostics.Stopwatch]::StartNew()
    $start    = [Diagnostics.Stopwatch]::StartNew()

    while ($true) {
        if ($script:Cancel) { return $false }

        if (Wait-Frame -Types @('DONE') -Session $Session -TimeoutSec 2) {
            $SawDone.Value = $true
            return $true
        }

        $len = 0L
        if (Test-Path -LiteralPath $Path) { $len = (Get-Item -LiteralPath $Path).Length }

        if ($len -ne $lastLen) {
            $lastLen = $len
            $lastMove.Restart()
            $rate = $len / [Math]::Max(0.001, $start.Elapsed.TotalSeconds)
            $eta  = if ($rate -gt 0) { [TimeSpan]::FromSeconds([Math]::Max(0L, $Size - $len) / $rate).ToString('hh\:mm\:ss') } else { '--' }
            Show-Advance $Name $len $Size $rate 0 $eta
        } elseif ($lastMove.Elapsed.TotalSeconds -gt $StallTimeoutSec) {
            Say '  the direct copy stopped making progress' 'warn'
            return $false
        }
    }
}

# Whatever carried the bytes, this is the only place a file becomes real: hash
# it, move it into place if it matches, and tell the sender either way.
function Complete-Incoming {
    param([string]$Source, [string]$Dest, [string]$Tmp, [string]$Part,
          [string]$FileHash, [string]$Session, [int64]$Size)

    if (-not (Test-Path -LiteralPath $Source)) {
        Say '  the transfer left nothing on disk' 'bad'
        Publish-Final -Frame (New-Frame -Type 'FIN' -Session $Session -Flags 'bad') -Session $Session
        return $false
    }

    $have = (Get-Item -LiteralPath $Source).Length
    if ($have -ne $Size) {
        Say "  short file - got $(Format-Size $have) of $(Format-Size $Size)" 'bad'
        Publish-Final -Frame (New-Frame -Type 'FIN' -Session $Session -Flags 'bad') -Session $Session
        return $false
    }

    Say '  verifying...' 'dim'
    $actual = Get-FileHashHex $Source

    if ($actual -eq $FileHash) {
        if (Test-Path -LiteralPath $Dest) { Remove-Item -LiteralPath $Dest -Force }
        Move-Item -LiteralPath $Source -Destination $Dest -Force
        Remove-Item -LiteralPath $Part -Force -ErrorAction SilentlyContinue
        if ($Source -ne $Tmp) { Remove-Item -LiteralPath $Tmp -Force -ErrorAction SilentlyContinue }
        Say ''
        Say "  saved to $Dest" 'ok'
        Publish-Final -Frame (New-Frame -Type 'FIN' -Session $Session -Flags 'ok') -Session $Session
        return $true
    }

    Say '  hash mismatch - file not saved' 'bad'
    Publish-Final -Frame (New-Frame -Type 'FIN' -Session $Session -Flags 'bad') -Session $Session
    return $false
}

# What to leave on the clipboard between transport attempts. Once the file is
# whole it has to stay an ACK - the sender is waiting to hear its last chunk
# landed, and a READY costs it a timeout and a chunk-size backoff.
function Get-NextOutFrame {
    param([string]$Session, [int64]$At, [int64]$Size, [string]$Advert)
    if ($At -ge $Size) { return New-Frame -Type 'ACK' -Session $Session -A $At }
    return New-Frame -Type 'READY' -Session $Session -A $At -Payload $Advert
}

# Receive one offered file: advertise what this side can do, serve whichever
# route the sender picks, and verify at the end.
function Receive-OneFile {
    param($Offer, [string]$Folder)

    $meta = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Offer.Payload)).Split("`n")
    $name = $meta[0]; $size = [int64]$meta[1]; $fileHash = $meta[2]
    $session = $Offer.Session

    # Must be absolute: the other side turns this into a \\tsclient\ path.
    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $Folder = (Resolve-Path -LiteralPath $Folder).Path

    $dest = Join-Path $Folder $name
    $tmp  = "$dest.sisitmp"
    $part = "$dest.sisipart"

    # Resume only if the sidecar is for this exact file.
    $offset = 0L
    if ((Test-Path -LiteralPath $part) -and (Test-Path -LiteralPath $tmp)) {
        $s = Get-Content -LiteralPath $part -Raw -ErrorAction SilentlyContinue
        if ($s) {
            $sp = $s.Split("`n")
            if ($sp[0] -eq $fileHash) { $offset = [int64]$sp[1] }
        }
    }

    Say ''
    Say "Incoming: $name  ($(Format-Size $size))" 'info'
    if ($offset -gt 0) { Say "  resuming at $(Format-Size $offset)" 'warn' }

    # Up before we answer, so the advert can name a port that is already open.
    $listener = Start-DataListener -BasePort $TcpPort

    $fs = [IO.File]::Open($tmp, 'OpenOrCreate', 'Write')
    $fsOpen = $true
    $ok = $false
    $sawDone = $false

    $verify = $tmp

    try {
        $fs.SetLength($size)          # preallocate, avoids fragmenting a 3 GB file
        $fs.Seek($offset, 'Begin') | Out-Null

        $advert = New-Advert -Listener $listener -Folder $Folder
        $ready  = New-Frame -Type 'READY' -Session $session -A $offset -Payload $advert
        $outFrame = $ready

        while ($true) {
            if ($script:Cancel) { Say '  cancelled' 'warn'; break }

            # Checked here, not after the wait: a transport can come back having
            # already seen DONE for itself.
            if ($sawDone) {
                if ($fsOpen) { $fs.Dispose(); $fsOpen = $false }
                $ok = Complete-Incoming -Source $verify -Dest $dest -Tmp $tmp -Part $part `
                                        -FileHash $fileHash -Session $session -Size $size
                break
            }

            $f = Send-AndWait -Frame $outFrame -Expect @('USE','DATA','DONE') -Session $session `
                              -TimeoutSec $StallTimeoutSec -RetryEverySec 2
            if (-not $f) { Say '  stalled - giving up' 'bad'; break }

            if ($f.Type -eq 'DONE') { $sawDone = $true; continue }

            $ref = [ref]$offset

            # A push deletes our staging file, so anything tried after one has to
            # start over from nothing.
            if (-not $fsOpen) {
                Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
                $offset = 0L
                $ref = [ref]$offset
                $verify = $tmp
                $fs = [IO.File]::Open($tmp, 'OpenOrCreate', 'Write')
                $fsOpen = $true
                $fs.SetLength($size)
                $fs.Seek(0, 'Begin') | Out-Null
            }

            if ($f.Type -eq 'DATA') {
                Set-Transport 'clipboard'
                Set-ClipText (New-Frame -Type 'NAK' -Session $session -A $offset) | Out-Null
                Receive-ViaClipboard -Out $fs -Session $session -Size $size -Name $name `
                                     -PartPath $part -FileHash $fileHash -Offset $ref `
                                     -SawDone ([ref]$sawDone) | Out-Null
                $offset = $ref.Value
                $outFrame = Get-NextOutFrame -Session $session -At $offset -Size $size -Advert $advert
                continue
            }

            $kind = $f.Flags
            $detail = ''
            if ($f.Payload) {
                try { $detail = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($f.Payload)) } catch {}
            }

            $setOk = New-Frame -Type 'SET' -Session $session -Flags 'ok' -A $offset
            $setNo = New-Frame -Type 'SET' -Session $session -Flags 'no' -A $offset
            $moved = $false

            # if/elseif, not switch: 'continue' inside a switch continues the
            # switch rather than this loop.
            if ($kind -eq 'tcp') {
                if (-not $listener) { $outFrame = $setNo; continue }
                Set-Transport 'direct connection'
                Set-ClipText $setOk | Out-Null
                $moved = Receive-ViaTcp -Listener $listener -Session $session -Out $fs -Size $size `
                                        -Name $name -PartPath $part -FileHash $fileHash -Offset $ref

            } elseif ($kind -eq 'tspush') {
                Set-Transport 'direct drive copy'
                Set-ClipText $setOk | Out-Null
                if ($fsOpen) { $fs.Dispose(); $fsOpen = $false }
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                $verify = Join-Path $Folder (Get-DropName $name $fileHash)
                $moved = Wait-ForPush -Path $verify -Size $size -Name $name `
                                      -Session $session -SawDone ([ref]$sawDone)

            } elseif ($kind -eq 'tspull') {
                $src = Convert-ToTsclientPath $detail
                if (-not $src -or -not (Test-Path -LiteralPath $src)) { $outFrame = $setNo; continue }
                Set-Transport 'direct drive copy'
                Set-ClipText $setOk | Out-Null
                $moved = Receive-ViaTsclient -RemotePath $detail -Out $fs -Size $size -Name $name `
                                             -PartPath $part -FileHash $fileHash -Offset $ref `
                                             -Session $session

            } elseif ($kind -eq 'clip') {
                Set-Transport 'clipboard'
                Set-ClipText $setOk | Out-Null
                $moved = Receive-ViaClipboard -Out $fs -Session $session -Size $size -Name $name `
                                              -PartPath $part -FileHash $fileHash -Offset $ref `
                                              -SawDone ([ref]$sawDone)

            } else {
                $outFrame = $setNo
                continue
            }

            $offset = $ref.Value

            if (-not $moved -and -not $script:Cancel) {
                Say '  that channel did not work - waiting for the sender to try another' 'warn'
            }

            $outFrame = Get-NextOutFrame -Session $session -At $offset -Size $size -Advert $advert
        }
    } finally {
        if ($fsOpen) { try { $fs.Dispose() } catch {} }
        Stop-DataListener $listener
        Set-Transport ''
    }
    return $ok
}

function Receive-Files {
    param([string]$Folder)

    if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $Folder = (Resolve-Path $Folder).Path

    Say ''
    Say "MailSISIbox receiving into $Folder" 'info'
    Say 'Waiting for a file. Ctrl+C to stop.' 'dim'

    while ($true) {
        $offer = Wait-ForOffer 86400
        if ($offer) {
            Receive-OneFile -Offer $offer -Folder $Folder | Out-Null
            Say ''
            Say 'Waiting for the next file. Ctrl+C to stop.' 'dim'
        }
    }
}

#endregion

#region ui

# Every control a handler touches is at $script: scope, and no handler gets
# .GetNewClosure(). A closure runs in its own module scope, where $script:
# resolves to that module and not to this file - so it reads $script:Queue as
# $null. Handlers fire from the message loop, where that bites.
function Show-MainWindow {
    # Otherwise a high-DPI screen bitmap-scales the window and blurs it. Has to
    # happen before any form exists.
    try { [MailSISIbox.Native]::SetProcessDPIAware() | Out-Null } catch {}
    [Windows.Forms.Application]::EnableVisualStyles()

    # Having claimed DPI awareness we own the scaling: every coordinate below is
    # in 96-dpi units and gets multiplied up. Fonts scale themselves.
    $gdc = [Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    $k   = $gdc.DpiX / 96.0
    $gdc.Dispose()
    function Px { param([double]$n) return [int][Math]::Round($n * $k) }

    $script:Ink    = [Drawing.Color]::FromArgb(31, 35, 40)
    $muted  = [Drawing.Color]::FromArgb(107, 114, 128)
    $bg     = [Drawing.Color]::FromArgb(252, 252, 253)
    $panel  = [Drawing.Color]::FromArgb(245, 246, 248)
    $script:Edge   = [Drawing.Color]::FromArgb(214, 219, 226)

    $script:Form = New-Object Windows.Forms.Form
    $script:Form.Text = 'MailSISIbox'
    $script:Form.ClientSize = New-Object Drawing.Size((Px 604), (Px 500))
    $script:Form.MinimumSize = New-Object Drawing.Size((Px 520), (Px 460))
    $script:Form.BackColor = $bg
    $script:Form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $script:Form.StartPosition = 'CenterScreen'

    $title = New-Object Windows.Forms.Label
    $title.Text = 'MailSISIbox'
    $title.Font = New-Object Drawing.Font('Segoe UI', 15, [Drawing.FontStyle]::Bold)
    $title.ForeColor = $script:Ink
    $title.SetBounds((Px 22), (Px 16), (Px 300), (Px 30))
    $script:Form.Controls.Add($title)

    $script:Sub = New-Object Windows.Forms.Label
    $script:IdleSub = 'Keep this open on both machines.'
    $script:Sub.Text = $script:IdleSub
    $script:Sub.ForeColor = $muted
    $script:Sub.SetBounds((Px 24), (Px 46), (Px 560), (Px 20))
    $script:Form.Controls.Add($script:Sub)

    $script:Drop = New-Object Windows.Forms.Panel
    $script:Drop.SetBounds((Px 22), (Px 76), (Px 560), (Px 120))
    $script:Drop.BackColor = $panel
    $script:Drop.AllowDrop = $true
    $script:Drop.Anchor = 'Top,Left,Right'
    $script:Form.Controls.Add($script:Drop)

    $dropLabel = New-Object Windows.Forms.Label
    $dropLabel.Text = "Drop files here to send`r`n(or click to browse)"
    $dropLabel.TextAlign = 'MiddleCenter'
    $dropLabel.ForeColor = $muted
    $dropLabel.Font = New-Object Drawing.Font('Segoe UI', 10)
    $dropLabel.Dock = 'Fill'
    $dropLabel.AllowDrop = $true
    $script:Drop.Controls.Add($dropLabel)

    # The label fills the panel, so the border is painted on the label. Px is a
    # function the handler cannot reach, hence the width worked out here.
    $script:Dash = [single](Px 1.5)
    $dropLabel.Add_Paint({
        param($s, $e)
        $p = New-Object Drawing.Pen($script:Edge, $script:Dash)
        $p.DashStyle = [Drawing.Drawing2D.DashStyle]::Dash
        $e.Graphics.DrawRectangle($p, 0, 0, $s.Width - 1, $s.Height - 1)
        $p.Dispose()
    })

    $destLabel = New-Object Windows.Forms.Label
    $destLabel.Text = 'Save incoming files to'
    $destLabel.ForeColor = $muted
    $destLabel.SetBounds((Px 24), (Px 208), (Px 200), (Px 18))
    $script:Form.Controls.Add($destLabel)

    $script:DestBox = New-Object Windows.Forms.TextBox
    $script:DestBox.Text = (Join-Path $env:USERPROFILE 'Downloads')
    $script:DestBox.SetBounds((Px 22), (Px 228), (Px 460), (Px 26))
    $script:DestBox.BorderStyle = 'FixedSingle'
    $script:DestBox.Anchor = 'Top,Left,Right'
    $script:Form.Controls.Add($script:DestBox)

    $browse = New-Object Windows.Forms.Button
    $browse.Text = 'Browse'
    $browse.SetBounds((Px 492), (Px 227), (Px 90), (Px 28))
    $browse.FlatStyle = 'Flat'
    $browse.FlatAppearance.BorderColor = $script:Edge
    $browse.BackColor = [Drawing.Color]::White
    $browse.Anchor = 'Top,Right'
    $script:Form.Controls.Add($browse)

    $script:Status = New-Object Windows.Forms.Label
    $script:Status.Text = 'Listening for incoming files...'
    $script:Status.ForeColor = $script:Ink
    $script:Status.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
    $script:Status.SetBounds((Px 24), (Px 272), (Px 560), (Px 20))
    $script:Status.Anchor = 'Top,Left,Right'
    $script:Form.Controls.Add($script:Status)

    $script:Bar = New-Object Windows.Forms.ProgressBar
    $script:Bar.SetBounds((Px 22), (Px 296), (Px 560), (Px 10))
    $script:Bar.Style = 'Continuous'
    $script:Bar.Maximum = 1000
    $script:Bar.Anchor = 'Top,Left,Right'
    $script:Form.Controls.Add($script:Bar)

    $script:Detail = New-Object Windows.Forms.Label
    $script:Detail.Text = ''
    $script:Detail.ForeColor = $muted
    $script:Detail.SetBounds((Px 24), (Px 312), (Px 560), (Px 18))
    $script:Detail.Anchor = 'Top,Left,Right'
    $script:Form.Controls.Add($script:Detail)

    $script:Log = New-Object Windows.Forms.ListBox
    $script:Log.SetBounds((Px 22), (Px 340), (Px 560), (Px 140))
    $script:Log.BorderStyle = 'FixedSingle'
    $script:Log.BackColor = [Drawing.Color]::White
    $script:Log.ForeColor = $muted
    $script:Log.Anchor = 'Top,Bottom,Left,Right'
    $script:Form.Controls.Add($script:Log)

    $script:OnStatus = {
        param($Text, $Level)
        if ([string]::IsNullOrWhiteSpace($Text)) { return }
        $t = $Text.Trim()
        $script:Log.Items.Add(('{0:HH:mm:ss}  {1}' -f (Get-Date), $t)) | Out-Null
        $script:Log.TopIndex = $script:Log.Items.Count - 1
        $script:Status.Text = $t
        $script:Status.ForeColor = switch ($Level) {
            'ok'   { [Drawing.Color]::FromArgb(22, 128, 74) }
            'bad'  { [Drawing.Color]::FromArgb(190, 50, 50) }
            'warn' { [Drawing.Color]::FromArgb(170, 110, 20) }
            default { $script:Ink }
        }
    }

    $script:OnProgress = {
        param($Label, $Done, $Total, $Rate, $Chunk, $Eta)
        $script:Bar.Value = [Math]::Min(1000, [int](($Done / [Math]::Max(1L, [int64]$Total)) * 1000))
        $script:Status.Text = $Label
        if ($Rate -gt 0) {
            $script:Detail.Text = ('{0} of {1}   {2}/s   chunk {3}   eta {4}' -f `
                (Format-Size $Done), (Format-Size $Total), (Format-Size $Rate), (Format-Size $Chunk), $Eta)
        } else {
            $script:Detail.Text = ('{0} of {1}' -f (Format-Size $Done), (Format-Size $Total))
        }
        $script:Sub.Text = if ($script:TransportLabel) {
            if ($script:TransportDetail) { 'Moving over {0} - {1}' -f $script:TransportLabel, $script:TransportDetail }
            else { 'Moving over {0}' -f $script:TransportLabel }
        } else { $script:IdleSub }
    }

    $script:UiPump = { [Windows.Forms.Application]::DoEvents() }

    $script:Queue = New-Object Collections.Generic.Queue[string]
    $script:Busy  = $false

    $script:Enqueue = {
        param($paths)
        foreach ($p in $paths) {
            if (Test-Path -LiteralPath $p -PathType Leaf) { $script:Queue.Enqueue($p) }
        }
    }

    $dropHandler = {
        param($s, $e)
        if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [Windows.Forms.DragDropEffects]::Copy
        }
    }
    $script:Drop.Add_DragEnter($dropHandler)
    $dropLabel.Add_DragEnter($dropHandler)

    $dropped = {
        param($s, $e)
        & $script:Enqueue $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    }
    $script:Drop.Add_DragDrop($dropped)
    $dropLabel.Add_DragDrop($dropped)

    $pick = {
        $d = New-Object Windows.Forms.OpenFileDialog
        $d.Multiselect = $true
        if ($d.ShowDialog() -eq 'OK') { & $script:Enqueue $d.FileNames }
    }
    $dropLabel.Add_Click($pick)

    $browse.Add_Click({
        $d = New-Object Windows.Forms.FolderBrowserDialog
        if ($d.ShowDialog() -eq 'OK') { $script:DestBox.Text = $d.SelectedPath }
    })

    $script:Timer = New-Object Windows.Forms.Timer
    $script:Timer.Interval = 150
    $script:Timer.Add_Tick({
        if ($script:Busy) { return }
        $script:Busy = $true
        try {
            if ($script:Queue.Count -gt 0) {
                $path = $script:Queue.Dequeue()
                Send-OneFile -Path $path | Out-Null
                $script:Bar.Value = 0; $script:Detail.Text = ''; $script:Sub.Text = $script:IdleSub
                Say 'Listening for incoming files...' 'dim'
            } else {
                $folder = $script:DestBox.Text
                if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
                $offer = Wait-ForOffer 0.4
                if ($offer) {
                    Receive-OneFile -Offer $offer -Folder $folder | Out-Null
                    $script:Bar.Value = 0; $script:Detail.Text = ''; $script:Sub.Text = $script:IdleSub
                    Say 'Listening for incoming files...' 'dim'
                }
            }
        } catch {
            Say ("error: " + $_.Exception.Message) 'bad'
        } finally {
            $script:Busy = $false
        }
    })
    $script:Timer.Start()

    if ($Minimized) { $script:Form.WindowState = [Windows.Forms.FormWindowState]::Minimized }

    # Under -WindowStyle Hidden, which is how the login entry and the launcher
    # start it, Windows applies that show-state to this form too and it never
    # appears. Force the state that was actually asked for.
    $script:Form.Add_Shown({
        $h = $script:Form.Handle
        if ($Minimized) {
            [void][MailSISIbox.Native]::ShowWindow($h, 7)      # minimised, no focus
        } else {
            [void][MailSISIbox.Native]::ShowWindow($h, 5)
            [void][MailSISIbox.Native]::SetForegroundWindow($h)
        }
        $script:DestBox.SelectionLength = 0
        $script:Drop.Focus() | Out-Null
    })
    $script:Form.Add_FormClosing({ $script:Cancel = $true; $script:Timer.Stop() })

    [void]$script:Form.ShowDialog()
}

#endregion

#region one window per machine

# Two windows on one machine both answer the far end's offers: one wins, the
# other holds an empty file it never receives and reports a hash mismatch, which
# reads as a connection fault. Console -Send / -Receive stay unguarded, since
# running both ends on one machine is how this gets tested.
function Enter-SingleInstance {
    $script:InstanceLock = New-Object Threading.Mutex($false, "Local\MailSISIbox-$env:USERNAME")
    try { return $script:InstanceLock.WaitOne(0) }
    catch [Threading.AbandonedMutexException] { return $true }   # last copy died holding it
}

#endregion

$script:LastSeq = [MailSISIbox.Native]::GetClipboardSequenceNumber()

if ($Receive) {
    Receive-Files -Folder $Receive
} elseif ($Send) {
    $saved = Get-ClipText
    try {
        foreach ($f in $Send) {
            foreach ($item in (Get-ChildItem -LiteralPath $f -File)) {
                Send-OneFile -Path $item.FullName | Out-Null
            }
        }
    } finally {
        Set-ClipText $saved | Out-Null
    }
    Write-Host ''
} elseif (Enter-SingleInstance) {
    Show-MainWindow
} elseif (-not $Minimized) {
    # Started by hand while another copy is up. Silent when it is the login copy
    # arriving second, since nobody asked for it.
    [void][Windows.Forms.MessageBox]::Show(
        "MailSISIbox is already running on this machine.`r`n`r`nLook for its window in the taskbar. Running two at once stops transfers working.",
        'MailSISIbox', 'OK', 'Information')
}
