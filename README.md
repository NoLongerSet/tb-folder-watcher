# FolderWatcher

A [twinBASIC](https://twinbasic.com/) utility that monitors a directory for new files and calls a VBA function in a running Microsoft Access application when one appears.

Built for the [Access DevCon Vienna 2026](https://www.accessdevcon.com/) talk on leveraging twinBASIC from Access applications.

## Why?

Access VBA has no way to receive file system notifications. The only option is polling with a `Form_Timer` event and `Dir()` calls — slow, CPU-wasteful, and it blocks the UI.

FolderWatcher uses Windows APIs that are simply unavailable from VBA to deliver a better experience:

| Capability | VBA Polling | FolderWatcher |
|---|---|---|
| **Detection method** | `Dir()` in a timer loop | `ReadDirectoryChangesW` (OS-level) |
| **Response time** | Seconds (depends on timer interval) | Instant |
| **CPU usage while idle** | Continuous (timer fires repeatedly) | Zero (thread sleeps in kernel) |
| **Knows the filename** | Must diff directory listings | Exact filename from the OS |
| **Blocks Access UI** | Yes, during each `Dir()` scan | No — runs in a separate process |
| **Cleanup on exit** | Manual (must remember to stop timer) | Automatic (detects Access closing) |

## Features

### Instant File Notifications

Uses the Windows `ReadDirectoryChangesW` API with overlapped I/O to receive file system events directly from the OS kernel. When a new file appears in the watched directory, the callback fires in your Access application within milliseconds — no polling, no timer intervals.

### Zero Dependencies

Single `.exe` file. No DLL registration, no COM components to install, no runtime prerequisites. Place it next to your Access database and it just works. Builds are available for both 32-bit and 64-bit Access.

### On-Demand COM Callbacks

When a new file is detected, FolderWatcher connects to your running Access instance via `GetObject()`, calls `Application.Run` with the file path, and immediately releases the COM reference. The reference is held for only milliseconds — never long enough to interfere with Access.

Your callback function can do anything: show a notification, import the file, log to a table, move the file to an archive folder, or kick off a workflow.

### Automatic Shutdown — No Orphan Processes

FolderWatcher opens a handle to the Access process with `SYNCHRONIZE` rights. When Access closes, the OS kernel signals this handle. The main loop uses `WaitForMultipleObjects` to wait on both the directory change event and the process handle simultaneously, so when Access exits, the watcher wakes up and shuts down instantly.

This is not polling. The watcher thread is fully asleep in the kernel until either a file event or a process exit occurs — zero CPU, zero overhead, instant response to both.

A 5-minute fallback timeout provides a belt-and-suspenders safety net in case the process handle can't be obtained.

### Bitness-Aware

The sample VBA module uses `#If Win64` conditional compilation to select the correct executable (`FolderWatcher_win32.exe` or `FolderWatcher_win64.exe`) to match your Access installation.

## Usage

```
FolderWatcher.exe --dir "C:\Inbox" --db "C:\MyApp.accdb" --function OnNewFile --pid 12345
```

| Flag | Required | Description |
|------|----------|-------------|
| `--dir` | Yes | Directory to watch for new files |
| `--db` | Yes | Full path to the running Access database (for `GetObject`) |
| `--function` | Yes | Public VBA function to call (must accept one `String` parameter) |
| `--pid` | Yes | Access process ID (watcher exits when this process closes) |

## Quick Start

### 1. Add the callback function to your Access database

Create a **standard module** with a public function that accepts a file path:

```vba
Public Function OnNewFile(ByVal FilePath As String) As Boolean
    Debug.Print "New file detected: " & FilePath
    MsgBox "New file detected:" & vbCrLf & vbCrLf & FilePath, vbInformation, "Folder Watcher"
    OnNewFile = True
End Function
```

### 2. Launch the watcher from VBA

```vba
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Sub StartWatching()
    Dim exePath As String
    #If Win64 Then
    exePath = CurrentProject.Path & "\FolderWatcher_win64.exe"
    #Else
    exePath = CurrentProject.Path & "\FolderWatcher_win32.exe"
    #End If

    Dim cmd As String
    cmd = """" & exePath & """" & _
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

See [`samples/modFolderWatcher.bas`](samples/modFolderWatcher.bas) for a complete Access VBA module with `StartWatching`, `StopWatching`, and a sample `OnNewFile` callback. A sample Access database source is also included in [`samples/FolderWatcherSample.accdb.src/`](samples/FolderWatcherSample.accdb.src/).

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

The main loop in [`Main.twin`](Source/Sources/App/Main.twin) waits on two kernel objects simultaneously:

1. **Directory change event** — signaled by `ReadDirectoryChangesW` when a file is added
2. **Access process handle** — signaled by the OS when the Access process exits

`WaitForMultipleObjects` blocks the thread until one of these fires. No CPU is consumed while waiting. When a file event arrives, it parses the `FILE_NOTIFY_INFORMATION` buffer to get the exact filename, makes the COM callback, and re-arms the watch. When the process handle fires, it exits cleanly.

## Building from Source

### Prerequisites

- [twinBASIC](https://twinbasic.com/) IDE (free)

### Steps

1. Open twinBASIC and create a new **Standard EXE** project
2. Import the `.twin` files from [`Source/Sources/`](Source/Sources/) into the project
3. Set the startup object to `Sub Main` in `EntryPoint`
4. Build for x86 (32-bit Access) and/or x64 (64-bit Access)

The [`Source/Settings`](Source/Settings) file contains the full project configuration (references, build options, compiler flags).

## Architecture

Follows the **functional core / imperative shell** pattern from the [NLS Launcher](https://github.com/NoLongerSet/DevCon2026) project:

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

| Layer | Responsibility | Rules |
|-------|---------------|-------|
| **Core** | Pure functions, data transformation | No Windows API calls, no I/O, no side effects |
| **Shell** | Windows API wrappers | File system, process management, COM automation |
| **App** | Orchestration | Wires Core and Shell together, owns the main loop |

## License

MIT
