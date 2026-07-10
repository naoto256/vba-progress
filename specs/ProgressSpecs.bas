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
    Spec_UpdateWithoutBeginFails
    Spec_ScopeWithoutBeginFails
    Spec_FinishWithoutBeginFails
    Spec_ResetRecoversFromActiveScope
    Spec_ProgressIsMonotonic

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

Public Sub Spec_UpdateWithoutBeginFails()
    Dim errorNumber As Long

    ResetProgressForSpec

    On Error Resume Next
    Progress.Update 0.5
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertErrorRaised errorNumber, "update without begin"

    ResetProgressForSpec
End Sub

Public Sub Spec_ScopeWithoutBeginFails()
    Dim guard As ProgressScopeGuard
    Dim errorNumber As Long

    ResetProgressForSpec

    On Error Resume Next
    Set guard = Progress.Scope(0, 0.5)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertErrorRaised errorNumber, "scope without begin"

    Set guard = Nothing
    ResetProgressForSpec
End Sub

Public Sub Spec_FinishWithoutBeginFails()
    Dim errorNumber As Long

    ResetProgressForSpec

    On Error Resume Next
    Progress.Finish
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertErrorRaised errorNumber, "finish without begin"

    ResetProgressForSpec
End Sub

Public Sub Spec_ResetRecoversFromActiveScope()
    Dim guard As ProgressScopeGuard
    Dim capture As ProgressSpecListener
    Dim errorNumber As Long

    ResetProgressForSpec
    Set capture = New ProgressSpecListener
    Progress.AddListener "capture", capture

    Progress.Begin
    Set guard = Progress.Scope(0, 0.5)

    ' Reset must recover even though a scope guard is still alive.
    Progress.Reset

    ' The normal lifecycle works again after recovery.
    capture.Clear
    Progress.Begin
    Progress.Finish
    AssertLongEqual 2, capture.Count, "recovered operation notification count"
    AssertNear 0, capture.Item(1), "recovered operation begin value"
    AssertNear 1, capture.Item(2), "recovered operation finish value"

    ' Destroying the now-stale guard must not leak an error: its Class_Terminate
    ' calls CloseScope, which fails on the discarded scope but swallows it.
    On Error Resume Next
    Set guard = Nothing
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If errorNumber <> 0 Then
        Err.Raise vbObjectError + 1003, "ProgressSpecs.stale guard terminate", _
            "Stale guard leaked an error."
    End If

    ResetProgressForSpec
End Sub

Public Sub Spec_ProgressIsMonotonic()
    Dim capture As ProgressSpecListener
    Dim countBeforeSecondScope As Long

    ResetProgressForSpec
    Set capture = New ProgressSpecListener
    Progress.AddListener "capture", capture

    Progress.Begin
    With Progress.Scope(0, 0.6)
        Progress.Update 1
        AssertNear 0.6, capture.LastValue, "first scope completes at 0.6"
    End With

    countBeforeSecondScope = capture.Count

    With Progress.Scope(0.6, 1)
        ' Opening this scope reports local 0 (global 0.6), which is not greater
        ' than the last notified value, so it must not emit a new notification.
        AssertLongEqual countBeforeSecondScope, capture.Count, _
            "monotonic guard suppresses regressive scope open"
    End With

    Progress.Finish
    ResetProgressForSpec
End Sub

Private Sub ResetProgressForSpec()
    Progress.RemoveListener "capture"
    Progress.RemoveListener "failing"
    Progress.Reset
End Sub

Private Sub AssertLongEqual(ByVal expected As Long, ByVal actual As Long, ByVal context As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 1000, "ProgressSpecs." & context, _
            "Expected " & CStr(expected) & ", got " & CStr(actual) & "."
    End If
End Sub

Private Sub AssertNear(ByVal expected As Double, ByVal actual As Double, ByVal context As String)
    If Abs(expected - actual) > EPSILON Then
        Err.Raise vbObjectError + 1001, "ProgressSpecs." & context, _
            "Expected " & CStr(expected) & ", got " & CStr(actual) & "."
    End If
End Sub

Private Sub AssertErrorRaised(ByVal errorNumber As Long, ByVal context As String)
    If errorNumber = 0 Then
        Err.Raise vbObjectError + 1002, "ProgressSpecs." & context, "Expected an error."
    End If
End Sub
