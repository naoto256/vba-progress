VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressFailingListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProgressListener

Private Count_ As Long

Private Sub Class_Initialize()
    Count_ = 0
End Sub

Public Property Get Count() As Long
    Count = Count_
End Property

Private Sub IProgressListener_OnProgressUpdated(ByVal progressValue As Double)
    Count_ = Count_ + 1
    Err.Raise vbObjectError + 2000, "ProgressFailingListener", "Intentional listener failure."
End Sub
