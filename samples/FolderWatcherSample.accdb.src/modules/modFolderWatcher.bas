Attribute VB_Name = "modFolderWatcher"
Option Compare Database
Option Explicit

Private Const APP_VERSION As String = "1.0.0"

' -----------------------------------------------------------------------
' modFolderWatcher
'
' Sample Access VBA module that demonstrates launching and using
' FolderWatcher.exe to receive callbacks when new files appear.
'
' The exe is embedded in the usys_Resources table and extracted
' automatically on first use.
'
' Usage:
'   1. (First time setup) From the Immediate Window, run:
'        ImportExe "C:\path\to\FolderWatcher_win32.exe"
'        ImportExe "C:\path\to\FolderWatcher_win64.exe"
'   2. Call StartWatching with a folder path -- the exe is extracted automatically
'   3. Drop files into the watched folder -- OnNewFile fires automatically
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

Private Const RESOURCE_TABLE As String = "usys_Resources"
Private Const RESOURCE_NAME_32 As String = "FolderWatcher_win32.exe"
Private Const RESOURCE_NAME_64 As String = "FolderWatcher_win64.exe"
Private Const APP_ICON As String = "folderwatcher.ico"

' Tracks the watcher's process ID so we can stop it later
Private m_watcherTaskId As Long


Public Function GetAppVersion() As String
    GetAppVersion = APP_VERSION
End Function

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

    Dim ExePath As String
    #If Win64 Then
    ExePath = CurrentProject.Path & "\FolderWatcher_win64.exe"
    #Else
    ExePath = CurrentProject.Path & "\FolderWatcher_win32.exe"
    #End If

    EnsureExeExists ExePath

    If Dir(ExePath) = "" Then
        MsgBox "FolderWatcher.exe not found in:" & vbCrLf & CurrentProject.Path, vbCritical
        Exit Sub
    End If

    Dim cmd As String
    cmd = """" & ExePath & """" & _
          " --dir """ & FolderPath & """" & _
          " --db """ & CurrentProject.FullName & """" & _
          " --function " & CallbackFunction & _
          " --pid " & GetCurrentProcessId()

    m_watcherTaskId = Shell(cmd, vbHide)

    If m_watcherTaskId = 0 Then
        MsgBox "Failed to start FolderWatcher.exe", vbCritical
    End If
End Sub

' Stops the folder watcher if it is running.
' Note: The watcher also stops automatically when Access closes.
Public Sub StopWatching()
    If m_watcherTaskId = 0 Then Exit Sub

    ' Shell() returns the process ID -- use it to terminate the watcher
    #If VBA7 Then
    Dim hProcess As LongPtr
    #Else
    Dim hProcess As Long
    #End If

    hProcess = OpenProcess(PROCESS_TERMINATE, 0, m_watcherTaskId)

    If hProcess <> 0 Then
        TerminateProcess hProcess, 0
        CloseHandle hProcess
    End If

    m_watcherTaskId = 0
End Sub

' -----------------------------------------------------------------------
' Private helpers
' -----------------------------------------------------------------------

' Extracts the app icon from usys_Resources and sets it as the Access
' application icon. Call from a startup form or AutoExec macro.
Public Sub SetAppIcon()
    Dim fpIcon As String
    fpIcon = CurrentProject.Path & "\" & APP_ICON

    ' Extract the .ico if it doesn't exist on disk yet
    If Dir(fpIcon) = "" Then
        If Not ExtractResource(APP_ICON, fpIcon) Then Exit Sub
    End If

    ' Set the AppIcon database property (creates it if needed)
    On Error Resume Next
    Dim prp As DAO.Property
    If CurrentDb.Properties("AppIcon") <> fpIcon Then
        CurrentDb.Properties("AppIcon") = fpIcon
    End If
    If Err.Number = 3270 Then
        ' Property doesn't exist yet -- create it
        Err.Clear
        Set prp = CurrentDb.CreateProperty("AppIcon", dbText, fpIcon)
        CurrentDb.Properties.Append prp
    End If
    On Error GoTo 0

    Application.RefreshTitleBar
End Sub

' Extracts the exe from the usys_Resources table if it doesn't already
' exist on disk next to the database.
Private Sub EnsureExeExists(ByVal ExePath As String)
    If Dir(ExePath) <> "" Then Exit Sub

    Dim ResourceName As String
    ResourceName = Mid$(ExePath, InStrRev(ExePath, "\") + 1)

    ExtractResource ResourceName, ExePath
End Sub

' Reads a named resource from usys_Resources and writes it to DestPath.
' Returns True on success, False on any failure (missing table, missing
' record, or disk write error).
Private Function ExtractResource(ByVal ResourceName As String, _
                                  ByVal DestPath As String) As Boolean
    If Not ResourceTableExists() Then Exit Function

    Dim db As DAO.Database
    Set db = CurrentDb
    With db.OpenRecordset(RESOURCE_TABLE, dbOpenSnapshot)
        .FindFirst "ResourceName=" & Qt(ResourceName)
        If .NoMatch Then Exit Function

        Dim b() As Byte
        b = !ResourceData.Value

        On Error GoTo WriteError
        Dim FNum As Integer
        FNum = FreeFile
        Open DestPath For Binary Lock Write As #FNum
        Put #FNum, , b
        Close #FNum
        On Error GoTo 0

        ExtractResource = True
        Exit Function

WriteError:
        Close #FNum
    End With
End Function

Private Function ResourceTableExists() As Boolean
    ResourceTableExists = (DCount("*", "MSysObjects", _
        "Name=" & Qt(RESOURCE_TABLE) & " AND Type=1") > 0)
End Function

' Simplified version of https://nolongerset.com/quoth-thy-sql-evermore/
Private Function Qt(ByVal s As String) As String
    Qt = Chr$(34) & s & Chr$(34)
End Function

' -----------------------------------------------------------------------
' Developer utilities -- run from the Immediate Window to populate the
' usys_Resources table with the FolderWatcher executables.
' -----------------------------------------------------------------------

' Creates the usys_Resources table if it does not already exist.
Public Sub CreateResourceTable()
    If ResourceTableExists() Then
        Debug.Print "Table already exists: " & RESOURCE_TABLE
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE " & RESOURCE_TABLE & " " & _
        "(ResourceName TEXT(255) NOT NULL PRIMARY KEY, " & _
        " ResourceData LONGBINARY NOT NULL)", dbFailOnError

    Debug.Print "Created table: " & RESOURCE_TABLE
End Sub

' Imports an exe (or any file) into the usys_Resources table.
' If a record with the same name already exists, it is overwritten.
'
' Usage:
'   ImportExe "C:\Build\FolderWatcher_win32.exe"
'   ImportExe "C:\Build\FolderWatcher_win64.exe"
Public Sub ImportExe(ByVal ExePath As String)
    If Dir(ExePath) = "" Then
        MsgBox "File not found: " & ExePath, vbCritical
        Exit Sub
    End If

    If Not ResourceTableExists() Then CreateResourceTable

    Dim ResourceName As String
    ResourceName = Mid$(ExePath, InStrRev(ExePath, "\") + 1)

    Dim FNum As Integer
    FNum = FreeFile
    Open ExePath For Binary Lock Read As #FNum
    Dim b() As Byte
    ReDim b(1 To LOF(FNum))
    Get #FNum, , b
    Close #FNum

    Dim db As DAO.Database
    Set db = CurrentDb
    With db.OpenRecordset(RESOURCE_TABLE, dbOpenDynaset)
        .FindFirst "ResourceName=" & Qt(ResourceName)
        If .NoMatch Then
            .AddNew
            !ResourceName.Value = ResourceName
        Else
            .Edit
        End If
        !ResourceData.Value = b
        .Update
        .Close
    End With

    Debug.Print "Imported: " & ResourceName & " (" & UBound(b) & " bytes)"
End Sub

Public Sub ReimportFolderWatcherExes()
    ImportExe CurrentProject.Path & "\..\Build\FolderWatcher_win32.exe"
    ImportExe CurrentProject.Path & "\..\Build\FolderWatcher_win64.exe"
    ImportExe CurrentProject.Path & "\folderwatcher.ico"
End Sub

' -----------------------------------------------------------------------
' Sample callback function -- called by FolderWatcher.exe via COM automation
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
