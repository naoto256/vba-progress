VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressSpecListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProgressListener

Private Values_ As Collection

Private Sub Class_Initialize()
    Set Values_ = New Collection
End Sub

Public Sub Clear()
    Set Values_ = New Collection
End Sub

Public Property Get Count() As Long
    Count = Values_.Count
End Property

Public Property Get Item(index As Long) As Double
    Item = CDbl(Values_(index))
End Property

Public Property Get LastValue() As Double
    If Values_.Count = 0 Then
        Err.Raise 5, "ProgressSpecListener.LastValue", "No progress values were recorded."
    End If

    LastValue = CDbl(Values_(Values_.Count))
End Property

Private Sub IProgressListener_OnProgressUpdated(progressValue As Double)
    Values_.Add progressValue
End Sub
