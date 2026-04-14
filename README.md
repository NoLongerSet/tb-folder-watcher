# FolderWatcher

A [twinBASIC](https://twinbasic.com/) utility that monitors a directory for new files and calls a VBA function in a running Microsoft Access application when one appears.

Built for the [Access DevCon Vienna 2026](https://www.accessdevcon.com/) talk on leveraging twinBASIC from Access applications.

## Why?

Access VBA's only option for detecting new files is polling with a `Form_Timer` event and `Dir()` calls — slow, CPU-wasteful, and it blocks the UI.

FolderWatcher runs as a separate process and uses Windows APIs to deliver a fundamentally better experience:

| Capability | VBA Polling | FolderWatcher |
|---|---|---|
| **Detection method** | `Dir()` in a timer loop | `ReadDirectoryChangesW` (OS-level) |
| **Response time** | Seconds (depends on timer interval) | Instant |
| **CPU usage while idle** | Continuous (timer fires repeatedly) | Zero (thread sleeps in kernel) |
| **Knows the filename** | Must diff directory listings | Exact filename from the OS |
| **Blocks Access UI** | Yes, during each `Dir()` scan | No — runs in a separate process |
| **Cleanup on exit** | Manual (must remember to stop timer) | Automatic (detects Access closing) |

## Why can't I simply call these APIs from Access?

You can — VBA can `Declare` and call every Win32 API that FolderWatcher uses. `ReadDirectoryChangesW`, `WaitForMultipleObjects`, `OpenProcess` — all of them work from VBA. The APIs aren't the problem. The execution model is.

**VBA is single-threaded and runs inside the Access process.** That one constraint makes the efficient approach unusable:

- **`WaitForMultipleObjects` blocks the calling thread.** In FolderWatcher's exe, that's fine — the thread has nothing else to do. In VBA, that's the *only* thread. Call it and Access freezes completely — no form interaction, no repainting, nothing — until a file appears or the timeout expires.

- **You can't move VBA code to a separate process.** There's no way to `Shell` a VBA script. VBA code runs inside `MSACCESS.EXE`, so there's nowhere to put a blocking wait that won't freeze the UI.

- **The VBA workaround is just polling with extra steps.** You *could* set up `ReadDirectoryChangesW` with overlapped I/O from VBA, then use `Form_Timer` to periodically call `GetOverlappedResult` with `bWait=False` to check if the event fired. But now you're polling on a timer again — you've replaced `Dir()` with a fancier notification buffer while losing the "zero CPU, instant response, thread sleeps in the kernel" benefit.

| | VBA (in-process) | twinBASIC .exe (separate process) |
|---|---|---|
| Can call `ReadDirectoryChangesW` | Yes | Yes |
| Can call `WaitForMultipleObjects` | Yes, but freezes Access | Yes — own process, no impact |
| Can block without consequences | No | Yes |
| Can monitor a process handle for exit | Must poll with a timer | Kernel wakes the thread instantly |

**The separate process is the key enabler.** It's what makes blocking waits free, keeps Access responsive, and lets the OS kernel do all the work at zero CPU cost. twinBASIC's role is that it compiles to a standalone `.exe` using syntax that Access developers already know — same `Declare` statements, same `Sub`/`Function` structure, same API calling conventions.

## Features

### Instant File Notifications

Uses the Windows `ReadDirectoryChangesW` API with overlapped I/O to receive file system events directly from the OS kernel. When a new file appears in the watched directory, the callback fires in your Access application within milliseconds — no polling, no timer intervals.

### Zero Dependencies

Single `.exe` file. No DLL registration, no COM components to install, no runtime prerequisites. Place it next to your Access database or [embed it directly in the `.accdb`](#embedding-the-executable) for fully self-contained distribution. Builds are available for both 32-bit and 64-bit Access.

### On-Demand COM Callbacks

When a new file is detected, FolderWatcher connects to your running Access instance via `GetObject()`, calls `Application.Run` with the file path, and immediately releases the COM reference. The reference is held for only milliseconds — never long enough to interfere with Access.

Your callback function can do anything: show a notification, import the file, log to a table, move the file to an archive folder, or kick off a workflow.

### Automatic Shutdown — No Orphan Processes

FolderWatcher opens a handle to the Access process with `SYNCHRONIZE` rights. When Access closes, the OS kernel signals this handle. The main loop uses `WaitForMultipleObjects` to wait on both the directory change event and the process handle simultaneously, so when Access exits, the watcher wakes up and shuts down instantly.

This is not polling. The watcher thread is fully asleep in the kernel until either a file event or a process exit occurs — zero CPU, zero overhead, instant response to both.

A 5-minute fallback timeout provides a belt-and-suspenders safety net in case the process handle can't be obtained.

### Bitness-Aware

The sample VBA module uses `#If Win64` conditional compilation to select the correct executable (`FolderWatcher_win32.exe` or `FolderWatcher_win64.exe`) to match your Access installation.

### Custom Application Icon

The sample database embeds a custom `.ico` file in the same `usys_Resources` table used for the executables. On startup, `SetAppIcon` extracts the icon next to the database and sets it as the Access application icon via the `AppIcon` database property. The icon appears in the Access title bar and taskbar — a small touch that makes the database feel like a polished, purpose-built application rather than a generic `.accdb` file.

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

## Embedding the Executable

Instead of distributing the `.exe` files alongside your database, you can embed them directly inside the `.accdb` as binary resources. The database becomes fully self-contained — one file to distribute, nothing to install.

This is a general-purpose technique for any small utility, but it's especially powerful when paired with [twinBASIC](https://twinbasic.com/). twinBASIC compiles to compact, dependency-free executables using syntax Access developers already know. That makes it practical to build purpose-built helper utilities — file watchers, background processors, system integrations — and ship them invisibly inside the database that uses them.

### How it works

The sample module stores executables in a `usys_Resources` table (the `usys_` prefix hides it from the Navigation Pane):

```sql
CREATE TABLE usys_Resources (
    ResourceName TEXT(255) NOT NULL PRIMARY KEY,
    ResourceData LONGBINARY NOT NULL
)
```

When `StartWatching` is called, it checks whether the `.exe` exists on disk next to the database. If not, it reads the raw bytes from `usys_Resources` and writes them to the file system before launching. On subsequent calls, the file is already there and extraction is skipped.

### Setup

From the Immediate Window, import the executables once:

```vba
ImportExe "C:\path\to\FolderWatcher_win32.exe"
ImportExe "C:\path\to\FolderWatcher_win64.exe"
```

This creates the `usys_Resources` table (if needed) and stores the binary contents. The `.accdb` is now self-deploying — distribute it to users and the exe is extracted automatically on first use.

After rebuilding the twinBASIC project, re-import with:

```vba
ReimportFolderWatcherExes
```

### Application icon

You can also embed a custom `.ico` in the same resource table:

```vba
ImportExe "C:\path\to\folderwatcher.ico"
```

The sample database calls `SetAppIcon` from the startup form's `Form_Open` event. This extracts the `.ico` next to the database (if not already there) and sets the `AppIcon` database property to the full path. Access picks up the icon in the title bar and taskbar after `Application.RefreshTitleBar`.

The `AppIcon` property requires a full absolute path — relative paths and the `rel:` prefix do not work. Because the path changes when the database moves to a different machine, `SetAppIcon` re-checks and updates the property on every open.

### Adapting the pattern

The resource table and extraction logic are generic. To embed a different file:

1. Import it: `ImportExe "C:\path\to\anything.dll"`
2. Extract it at runtime: call `ExtractResource "anything.dll", destPath`

The `LONGBINARY` field stores raw bytes with no OLE wrapper overhead, and the extraction code handles files of any size using simple binary I/O.

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

## Pre-built Binaries

Signed executables for both 32-bit and 64-bit Access are available on the [Releases](https://github.com/NoLongerSet/tb-folder-watcher/releases) page. Download the one that matches your Access installation and skip straight to the [Quick Start](#quick-start).

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
