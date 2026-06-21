VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Immutable value object describing how a child scope maps into its parent.
'
' `startPoint` and `endPoint` are expressed in the parent scope's local 0..1
' coordinate system. The manager composes these ranges across the active stack.
Private StartPoint_ As Double
Private EndPoint_ As Double

' Initializes an immutable mapping from local progress to a parent range.
'
' The range is inclusive and must move forward. Reverse ranges would make root
' progress non-monotonic and break listener assumptions.
Public Sub Init(startPoint As Double, endPoint As Double)
    If startPoint < 0 Or startPoint > 1 Then
        Err.Raise 5, "ProgressScope.Init", "Start point must be between 0 and 1."
    End If
    If endPoint < 0 Or endPoint > 1 Then
        Err.Raise 5, "ProgressScope.Init", "End point must be between 0 and 1."
    End If
    If endPoint < startPoint Then
        Err.Raise 5, "ProgressScope.Init", "End point must be greater than or equal to start point."
    End If

    StartPoint_ = startPoint
    EndPoint_ = endPoint
End Sub

' Inclusive start point in the parent scope.
Public Property Get startPoint() As Double
    startPoint = StartPoint_
End Property

' Inclusive end point in the parent scope.
Public Property Get endPoint() As Double
    endPoint = EndPoint_
End Property
