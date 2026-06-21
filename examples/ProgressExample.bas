Attribute VB_Name = "ProgressExample"
Option Explicit

Public Sub RunProgressExample()
    Progress.AddListener "status", New ProgressListener_AppStatusBar
    Progress.AddListener "debug", New ProgressListener_DebugPrint

    Progress.Begin

    With Progress.Scope(0, 0.3)
        SimulateWork 10
    End With

    With Progress.Scope(0.3, 0.8)
        SimulateWork 20
    End With

    With Progress.Scope(0.8, 1)
        SimulateWork 5
    End With

    Progress.Finish
End Sub

Private Sub SimulateWork(stepCount As Long)
    Dim i As Long

    For i = 1 To stepCount
        Progress.Update i / stepCount
        DoEvents
    Next i
End Sub
