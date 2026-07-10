VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressListener_AppStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProgressListener

' Excel status bar adapter.
'
' `DoEvents` keeps Excel responsive during long operations, but it can re-enter
' VBA code. `IsUpdating_` drops nested status-bar updates from that reentry.
Private IsUpdating_ As Boolean

Private Sub Class_Initialize()
    IsUpdating_ = False
End Sub

Private Sub IProgressListener_OnProgressUpdated(ByVal progressValue As Double)
    If IsUpdating_ Then
        Exit Sub
    End If

    IsUpdating_ = True
    On Error GoTo CleanUp

    If progressValue >= 1 Then
        Application.StatusBar = False
    Else
        Application.StatusBar = "Progress: " & Format(progressValue, "0.0%")
    End If
    
    DoEvents

    IsUpdating_ = False
    Exit Sub

CleanUp:
    IsUpdating_ = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
