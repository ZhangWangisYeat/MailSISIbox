# MailSISIbox - Version 1.0

Moves files between two Windows machines linked by a remote desktop connection,
for when drive redirection and file copy-paste are switched off. One PowerShell
script with no install, no admin rights, no third-party dependencies, no size limit.

## How it works

The text clipboard is the one channel that still works in a locked-down session,
so it is used as a control channel: the two sides find each other over it, then
negotiate the fastest transport they actually share.

| Route | When you get it |
|---|---|
| Direct TCP connection | the two machines can reach each other over the network |
| Drive copy over `\\tsclient\` | the session can see the other machine's drives |
| The clipboard itself | always — this is the fallback |

Files are sent in chunks, each hashed in transit, with a SHA-256 check over the
whole file before it is saved. Interrupted transfers resume where they stopped. 
Chunks are streamed rather than buffered, so a 3 GB file uses about as much memory
as a 50 MB one.

Chunk size, and whether compression pays at all, are measured against the live
link while it runs: round-trip time dominates over RDP, CPU dominates on a fast
local link, and the right answer differs by an order of magnitude between them.

## Running it

Run it on both machines, then drag a file onto either window to send it. Both
directions work from the same window.

```powershell
powershell -STA -ExecutionPolicy Bypass -File .\MailSISIbox.ps1
```

Only one copy per machine.

Without a window:

```powershell
.\MailSISIbox.ps1 -Receive C:\Drop      # wait for files, save them here
.\MailSISIbox.ps1 -Send C:\book.xlsx    # send one file
```

Options:

```
-InstallStartup / -RemoveStartup   start at login, or stop doing that
-Minimized                         open out of the way
-MaxChunkMB 16                     use less memory, at some cost to speed
-TcpPort 48731                     first port tried for a direct connection
-NoTcp / -NoTsclient               skip a route
-Trace                             per-chunk timings
```

## Requirements

Windows PowerShell 5.1, present on every Windows machine.

It also needs a desktop Remote Desktop client. The browser-based web client has
no clipboard channel a program can drive, as a browser only reaches the clipboard
when a human presses Ctrl+C or Ctrl+V, so the two sides can never find each
other through one.

# Notes

PowerShell MailSISIBox.cmd file not included here for security purposes. It contains 
the built-in interface within PowerShell to select and run commands without the need 
for the user to type any commands themselves. Lightweight version will be released
publicly in Version 1.1.
