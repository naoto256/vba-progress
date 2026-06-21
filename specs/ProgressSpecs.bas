Attribute VB_Name = "ProgressSpecs"
Option Explicit

Private Const EPSILON As Double = 0.000001

Public Sub RunAllProgressSpecs()
    On Error GoTo Fail

    Spec_RootOperationNotifiesBeginAndFinish
    Spec_NestedScopeMapsLocalProgress
    Spec_InvalidScopeRangeFails
    Spec_InvalidProgressValueFails
    Spec_FinishFailsInsideActiveScope
    Spec_ListenerFailureDoesNotStopOtherListeners

    Debug.Print "All Progress specs passed."
    Exit Sub

Fail:
    Debug.Print "Progress spec failed: " & Err.Source & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub Spec_RootOperationNotifiesBeginAndFinish()
    Dim capture As ProgressSpecListener

    ResetProgressForSpec
    Set capture = New ProgressSpecListener
    Progress.AddListener "capture", capture

    Progress.Begin
    Progress.Finish

    AssertLongEqual 2, capture.Count, "root operation notification count"
    AssertNear 0, capture.Item(1), "root operation begin value"
    AssertNear 1, capture.Item(2), "root operation finish value"

    ResetProgressForSpec
End Sub

Public Sub Spec_NestedScopeMapsLocalProgress()
    Dim capture As ProgressSpecListener

    ResetProgressForSpec
    Set capture = New ProgressSpecListener
    Progress.AddListener "capture", capture

    Progress.Begin
    With Progress.Scope(0.25, 0.75)
        Progress.Update 0.5
        AssertNear 0.5, capture.LastValue, "nested scope middle value"
    End With
    AssertNear 0.75, capture.LastValue, "nested scope close value"
    Progress.Finish

    ResetProgressForSpec
End Sub

Public Sub Spec_InvalidScopeRangeFails()
    Dim guard As ProgressScopeGuard
    Dim errorNumber As Long

    ResetProgressForSpec

    On Error Resume Next
    Set guard = Progress.Scope(0.8, 0.2)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertErrorRaised errorNumber, "invalid scope range"

    Set guard = Nothing
    ResetProgressForSpec
End Sub

Public Sub Spec_InvalidProgressValueFails()
    Dim errorNumber As Long

    ResetProgressForSpec
    Progress.Begin

    On Error Resume Next
    Progress.Update 1.1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertErrorRaised errorNumber, "invalid progress value"

    Progress.Finish
    ResetProgressForSpec
End Sub

Public Sub Spec_FinishFailsInsideActiveScope()
    Dim guard As ProgressScopeGuard
    Dim errorNumber As Long

    ResetProgressForSpec
    Progress.Begin
    Set guard = Progress.Scope(0, 0.5)

    On Error Resume Next
    Progress.Finish
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertErrorRaised errorNumber, "finish inside active scope"

    Set guard = Nothing
    Progress.Reset
    ResetProgressForSpec
End Sub

Public Sub Spec_ListenerFailureDoesNotStopOtherListeners()
    Dim capture As ProgressSpecListener
    Dim failing As ProgressFailingListener

    ResetProgressForSpec
    Set capture = New ProgressSpecListener
    Set failing = New ProgressFailingListener
    Progress.AddListener "failing", failing
    Progress.AddListener "capture", capture

    Progress.Begin
    Progress.Update 0.25

    AssertLongEqual 2, failing.Count, "failing listener call count"
    AssertLongEqual 2, capture.Count, "healthy listener call count"
    AssertNear 0.25, capture.LastValue, "healthy listener still receives update"

    Progress.Finish
    ResetProgressForSpec
End Sub

Private Sub ResetProgressForSpec()
    Progress.RemoveListener "capture"
    Progress.RemoveListener "failing"
    Progress.Reset
End Sub

Private Sub AssertLongEqual(expected As Long, actual As Long, context As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 1000, "ProgressSpecs." & context, _
            "Expected " & CStr(expected) & ", got " & CStr(actual) & "."
    End If
End Sub

Private Sub AssertNear(expected As Double, actual As Double, context As String)
    If Abs(expected - actual) > EPSILON Then
        Err.Raise vbObjectError + 1001, "ProgressSpecs." & context, _
            "Expected " & CStr(expected) & ", got " & CStr(actual) & "."
    End If
End Sub

Private Sub AssertErrorRaised(errorNumber As Long, context As String)
    If errorNumber = 0 Then
        Err.Raise vbObjectError + 1002, "ProgressSpecs." & context, "Expected an error."
    End If
End Sub
