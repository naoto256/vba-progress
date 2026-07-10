VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressListener_DebugPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProgressListener

' Immediate-window listener useful for examples and specs.
Private Sub IProgressListener_OnProgressUpdated(ByVal progressValue As Double)
    Debug.Print "Progress: " & Format(progressValue * 100, "0.0") & "%"
End Sub
