VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Holds normalized progress state for one root operation.
'
' Progress is represented as a stack of scopes. Each scope maps local 0..1
' progress into a range of its parent scope, so callers can compose nested work
' without manually recalculating global percentages.
Private ScopeStack_ As Collection
Private Listeners_ As Collection
Private LastProgressValue_ As Double
Private Active_ As Boolean
Private IsNotifying_ As Boolean

Private Sub Class_Initialize()
    Set Listeners_ = New Collection
    ResetState
End Sub

' Resets scopes and monotonic progress, but keeps registered listeners.
'
' Recovery entry point: unlike Begin / Finish, Reset never fails on a
' non-empty stack. It discards any residual scopes so callers can recover
' from a non-LIFO guard teardown or a half-finished operation. Live guards
' whose scopes are discarded will fail in Class_Terminate when they call
' CloseScope, but that failure is swallowed there, so this is safe.
Public Sub Reset()
    ResetState
End Sub

Private Sub ResetState()
    Set ScopeStack_ = New Collection
    LastProgressValue_ = -1
    Active_ = False

    Dim root As ProgressScope

    Set root = New ProgressScope
    root.Init 0, 1

    ScopeStack_.Add root
End Sub

' Starts a new root operation.
'
' A root operation always starts with a single 0..1 root scope.
Public Sub Begin()
    EnsureRootScope "ProgressManager.Begin"
    ResetState
    Active_ = True
    Update 0
End Sub

' Completes the root operation. Active child scopes must be closed first.
'
' Child scopes close themselves through `ProgressScopeGuard`; finishing from
' inside a child scope would make the guard's later termination ambiguous.
' Requires an active operation started with `Begin`.
Public Sub Finish()
    EnsureRootScope "ProgressManager.Finish"
    If Not Active_ Then
        Err.Raise 5, "ProgressManager.Finish", "No active operation. Call Begin first."
    End If
    Update 1
    ResetState
End Sub

' Registers or replaces a named listener.
'
' The replace-on-add behavior keeps setup code repeatable during development
' and in long-lived host projects.
Public Sub AddListener(ByVal name As String, ByVal listener As IProgressListener)
    RemoveListener name
    Listeners_.Add listener, name
End Sub

' Removes a named listener if it exists.
Public Sub RemoveListener(ByVal name As String)
    On Error Resume Next
    Listeners_.Remove name
    On Error GoTo 0
End Sub

' Internal entry point used by ProgressScopeGuard.
'
' This is `Friend` so it is hidden from external references, but still usable
' by other modules in the same VBAProject. VBA has no narrower "friend class"
' visibility.
Friend Sub PushScope(ByVal ps As ProgressScope)
    If ps Is Nothing Then
        Err.Raise 5, "ProgressManager.PushScope", "Progress scope is required."
    End If
    If Not Active_ Then
        Err.Raise 5, "ProgressManager.PushScope", "No active operation. Call Begin first."
    End If

    ScopeStack_.Add ps
End Sub

' Internal escape hatch for explicit scope pops in specs or advanced code.
'
' Normal callers should not use this directly. `CloseScope` is the usual guard
' path because it sends the final local 100% update before popping.
Friend Sub PopScope(ByVal ps As ProgressScope)
    EnsureTopScope ps, "ProgressManager.PopScope"
    ScopeStack_.Remove ScopeStack_.Count
End Sub

' Closes the active scope after notifying it at local 100%.
'
' The top-scope check is important: scope guards must be destroyed in LIFO
' order, matching how nested `With Progress.Scope(...)` blocks are written.
Friend Sub CloseScope(ByVal ps As ProgressScope)
    EnsureTopScope ps, "ProgressManager.CloseScope"
    Update 1
    ScopeStack_.Remove ScopeStack_.Count
End Sub

' Reports local progress in the active scope.
'
' Notifications are monotonic at the root level. This avoids noisy regressions
' when a child scope starts at local 0 after its parent has already advanced.
' Requires an active operation started with `Begin`.
Public Sub Update(ByVal localProgressValue As Double)
    If Not Active_ Then
        Err.Raise 5, "ProgressManager.Update", "No active operation. Call Begin first."
    End If
    ValidateProgressValue localProgressValue

    Dim progressValue As Double

    progressValue = CalculateProgressValue(localProgressValue)
    If progressValue > LastProgressValue_ Then
        LastProgressValue_ = progressValue
        NotifyListeners progressValue
    End If
End Sub

Private Function CalculateProgressValue(ByVal localProgressValue As Double) As Double
    Dim ps As ProgressScope

    Dim offset As Double
    Dim ratio As Double

    offset = 0
    ratio = 1

    ' Fold each nested scope into one affine transform:
    ' global = offset + ratio * localProgressValue
    For Each ps In ScopeStack_
        offset = offset + ratio * ps.startPoint
        ratio = ratio * (ps.endPoint - ps.startPoint)
    Next ps

    Dim result As Double
    result = offset + ratio * localProgressValue

    ' Folding scopes accumulates floating-point error, which can push the
    ' result a hair outside 0..1. Clamp so listeners never see, e.g., 1.0000001.
    If result < 0 Then
        result = 0
    ElseIf result > 1 Then
        result = 1
    End If

    CalculateProgressValue = result
End Function

Private Sub ValidateProgressValue(ByVal progressValue As Double)
    If progressValue < 0 Or progressValue > 1 Then
        Err.Raise 5, "ProgressManager.Update", "Progress value must be between 0 and 1."
    End If
End Sub

Private Sub EnsureRootScope(ByVal source As String)
    If ScopeStack_.Count <> 1 Then
        Err.Raise 5, source, "Cannot begin or finish while a scope is active."
    End If
End Sub

Private Sub EnsureTopScope(ByVal ps As ProgressScope, ByVal source As String)
    If ps Is Nothing Then
        Err.Raise 5, source, "Progress scope is required."
    End If
    If ScopeStack_.Count <= 1 Then
        Err.Raise 5, source, "No progress scope is active."
    End If
    If Not ScopeStack_(ScopeStack_.Count) Is ps Then
        Err.Raise 5, source, "Progress scope stack is out of order."
    End If
End Sub

Private Sub NotifyListeners(ByVal progressValue As Double)
    ' A listener may pump messages (e.g. `DoEvents`) and re-enter progress code.
    ' Guard against reentry so a nested notification cannot mutate Listeners_
    ' mid-iteration or clobber LastProgressValue_.
    If IsNotifying_ Then
        Exit Sub
    End If

    Dim listener As IProgressListener

    IsNotifying_ = True
    For Each listener In Listeners_
        ' Progress display is auxiliary. A broken UI/debug listener should not
        ' stop the operation or prevent other listeners from receiving updates.
        On Error Resume Next
        listener.OnProgressUpdated progressValue
        If Err.Number <> 0 Then
            Debug.Print "Progress listener failed: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next
    ' Always clear the flag, even if the loop exited early, so notifications
    ' are not permanently suppressed.
    IsNotifying_ = False
End Sub
