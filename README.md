# FolderWatcher

A twinBASIC utility that monitors a directory for new files and calls a VBA function in a running Microsoft Access application when one appears.

Built for the [Access DevCon Vienna 2026](https://www.accessdevcon.com/) talk on leveraging twinBASIC from Access applications.

## Why?

Access VBA has no way to receive file system notifications. The only option is polling with a `Form_Timer` event and `Dir()` calls — slow, CPU-wasteful, and blocks the UI.

FolderWatcher solves this with:

- **Instant notifications** via the Windows `ReadDirectoryChangesW` API
- **Zero dependencies** — single .exe, no DLL registration, no runtime to install
- **Clean COM integration** — calls back into your running Access app via `GetObject` + `Application.Run`
- **Self-cleaning** — automatically exits when Access closes (monitors the parent process)

## Usage

```
FolderWatcher.exe --dir "C:\Inbox" --db "C:\MyApp.accdb" --function OnNewFile --pid 12345
```

| Flag | Required | Description |
|------|----------|-------------|
| `--dir` | Yes | Directory to watch for new files |
| `--db` | Yes | Full path to the running Access database |
| `--function` | Yes | Public VBA function to call (must accept one `String` parameter) |
| `--pid` | Yes | Access process ID (watcher exits when this process closes) |

## Quick Start

### 1. Add the callback function to your Access database

Create a **standard module** with a public function that accepts a file path:

```vba
Public Function OnNewFile(ByVal FilePath As String) As Boolean
    MsgBox "New file detected: " & FilePath, vbInformation
    OnNewFile = True
End Function
```

### 2. Launch the watcher from VBA

```vba
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Sub StartWatching()
    Dim cmd As String
    cmd = """" & CurrentProject.Path & "\FolderWatcher.exe""" & _
          " --dir ""C:\Inbox""" & _
          " --db """ & CurrentProject.FullName & """" & _
          " --function OnNewFile" & _
          " --pid " & GetCurrentProcessId()

    Shell cmd, vbMinimizedNoFocus
End Sub
```

### 3. Drop a file into the watched folder

Your `OnNewFile` function fires immediately with the full path of the new file.

## Sample Module

See [`samples/modFolderWatcher.bas`](samples/modFolderWatcher.bas) for a complete Access VBA module with `StartWatching`, `StopWatching`, and a sample `OnNewFile` callback. Import it into your Access database to get started.

## How It Works

```
Access VBA                    FolderWatcher.exe                   Windows
-----------                   -------------------                 -------
Shell "FolderWatcher.exe"  ->  Parse args
                               OpenProcess(pid)               ->  Track Access PID
                               CreateFileW(dir)               ->  Open directory handle
                               ReadDirectoryChangesW()         ->  Register for notifications
                               WaitForMultipleObjects()        ->  Sleep (zero CPU)
                                        |
    [user drops file into folder]       |
                                        v
                               Parse FILE_NOTIFY_INFORMATION
                               GetObject(db) -+
                                              |
OnNewFile(filePath)  <---  Application.Run ---+
                               Set app = Nothing               Release COM reference
                               ReadDirectoryChangesW()         ->  Re-arm for next file
                                        |
    [user closes Access]                |
                                        v
                               WaitForMultipleObjects returns
                               (process handle signaled)
                               CloseHandle, exit cleanly
```

Key design decisions:
- **No persistent COM reference** to Access. `GetObject` connects on-demand for each callback and releases immediately, so Access can close cleanly at any time.
- **Process monitoring** via `WaitForMultipleObjects` on the Access process handle. The watcher exits instantly when Access closes — no orphan processes.
- **5-minute fallback poll** in case the process handle can't be obtained.

## Building from Source

1. Install [twinBASIC](https://twinbasic.com/) (free IDE)
2. Create a new **Standard EXE** project
3. Import the `.twin` files from [`Source/Sources/`](Source/Sources/) into the project
4. Set the startup object to `Sub Main` in `EntryPoint`
5. Build for x86 (matches 32-bit Access) or x64

The [`Source/Settings`](Source/Settings) file contains the project configuration (references, build options).

## Architecture

Follows the **functional core / imperative shell** pattern:

```
Source/Sources/
  Core/
    ArgumentParserCore.twin   Pure string-parsing helpers (no I/O)
  Shell/
    DirectoryWatcher.twin     ReadDirectoryChangesW + overlapped I/O
    ProcessMonitor.twin       OpenProcess / process lifetime tracking
    AccessCallback.twin       GetObject + Application.Run (on-demand COM)
  App/
    Main.twin                 Entry point + main watch loop
    ArgumentParser.twin       CLI argument parsing
```

- **Core** — Pure functions, no Windows API calls, no side effects
- **Shell** — Windows API wrappers (file system, process, COM)
- **App** — Orchestration that wires Core and Shell together

## License

MIT
