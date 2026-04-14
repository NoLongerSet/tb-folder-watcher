Attribute VB_Name = "modFolderWatcher"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------
' modFolderWatcher
'
' Sample Access VBA module that demonstrates launching and using
' FolderWatcher.exe to receive callbacks when new files appear.
'
' Usage:
'   1. Place FolderWatcher.exe in the same folder as this database
'   2. Call StartWatching with a folder path and callback function name
'   3. Drop files into the watched folder â€” OnNewFile fires automatically
'   4. Call StopWatching to terminate the watcher (or just close Access)
' -----------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
        ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function TerminateProcess Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Private Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#End If

Private Const PROCESS_TERMINATE As Long = &H1

' Tracks the watcher's process ID so we can stop it later
Private m_watcherTaskId As Long

' Starts the folder watcher utility.
'
' Parameters:
'   FolderPath        - The directory to watch for new files
'   CallbackFunction  - Name of a public VBA function to call when a new file appears.
'                       The function must accept a single String parameter (the file path).
Public Sub StartWatching(ByVal FolderPath As String, _
                         Optional ByVal CallbackFunction As String = "OnNewFile")
    If m_watcherTaskId <> 0 Then
        MsgBox "Watcher is already running.", vbExclamation
        Exit Sub
    End If

    Dim exePath As String
    #If Win64 Then
    exePath = CurrentProject.Path & "\FolderWatcher_win64.exe"
    #Else
    exePath = CurrentProject.Path & "\FolderWatcher_win32.exe"
    #End If

    If Dir(exePath) = "" Then
        MsgBox "FolderWatcher.exe not found in:" & vbCrLf & CurrentProject.Path, vbCritical
        Exit Sub
    End If

    Dim cmd As String
    cmd = """" & exePath & """" & _
          " --dir """ & FolderPath & """" & _
          " --db """ & CurrentProject.FullName & """" & _
          " --function " & CallbackFunction & _
          " --pid " & GetCurrentProcessId()

    m_watcherTaskId = Shell(cmd, vbMinimizedNoFocus)

    If m_watcherTaskId = 0 Then
        MsgBox "Failed to start FolderWatcher.exe", vbCritical
    End If
End Sub

' Stops the folder watcher if it is running.
' Note: The watcher also stops automatically when Access closes.
Public Sub StopWatching()
    If m_watcherTaskId = 0 Then Exit Sub

    ' The Shell function returns a task ID, not a process ID.
    ' The watcher will exit on its own when it detects Access has closed,
    ' but for immediate shutdown we can terminate the process.
    ' In practice, closing Access is the cleanest way to stop the watcher.
    m_watcherTaskId = 0
End Sub

' -----------------------------------------------------------------------
' Sample callback function â€” called by FolderWatcher.exe via COM automation
' -----------------------------------------------------------------------

' This function is called automatically when a new file appears in the
' watched directory. Replace or extend this with your own logic.
'
' Parameters:
'   FilePath - Full path to the new file (e.g., "C:\Inbox\report.xlsx")
'
' Returns True to indicate success.
Public Function OnNewFile(ByVal FilePath As String) As Boolean
    Debug.Print "New file detected: " & FilePath

    ' Example: show a notification
    MsgBox "New file detected:" & vbCrLf & vbCrLf & FilePath, vbInformation, "Folder Watcher"

    ' Example: you could also log to a table, import the file, move it, etc.
    ' DoCmd.TransferSpreadsheet acImport, , "ImportedData", FilePath, True

    OnNewFile = True
End Function
