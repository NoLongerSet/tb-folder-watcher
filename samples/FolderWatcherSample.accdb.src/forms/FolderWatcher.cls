Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    SetAppIcon
End Sub

Private Sub toggleWatchFolder_AfterUpdate()
    If Me.toggleWatchFolder.Value Then
        StartWatching Me.tbFolderToWatch.Value
        Me.toggleWatchFolder.Caption = "status: WATCHING" & vbNewLine & "(press to stop)"
    Else
        StopWatching
        Me.toggleWatchFolder.Caption = "status: NOT Watching" & vbNewLine & "(press to start)"
    End If
End Sub
