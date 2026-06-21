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

Private Sub Class_Initialize()
    Set Listeners_ = New Collection
    ResetState
End Sub

' Resets scopes and monotonic progress, but keeps registered listeners.
'
' Reset is only legal at the root. Resetting while a guard is alive would
' invalidate that guard's ownership of the stack top.
Public Sub Reset()
    EnsureRootScope "ProgressManager.Reset"
    ResetState
End Sub

Private Sub ResetState()
    Set ScopeStack_ = New Collection
    LastProgressValue_ = -1
    
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
    Update 0
End Sub

' Completes the root operation. Active child scopes must be closed first.
'
' Child scopes close themselves through `ProgressScopeGuard`; finishing from
' inside a child scope would make the guard's later termination ambiguous.
Public Sub Finish()
    EnsureRootScope "ProgressManager.Finish"
    Update 1
    ResetState
End Sub

' Registers or replaces a named listener.
'
' The replace-on-add behavior keeps setup code repeatable during development
' and in long-lived host projects.
Public Sub AddListener(name As String, listener As IProgressListener)
    RemoveListener name
    Listeners_.Add listener, name
End Sub

' Removes a named listener if it exists.
Public Sub RemoveListener(name As String)
    On Error Resume Next
    Listeners_.Remove name
    On Error GoTo 0
End Sub

' Internal entry point used by ProgressScopeGuard.
'
' This is `Friend` so it is hidden from external references, but still usable
' by other modules in the same VBAProject. VBA has no narrower "friend class"
' visibility.
Friend Sub PushScope(ps As ProgressScope)
    If ps Is Nothing Then
        Err.Raise 5, "ProgressManager.PushScope", "Progress scope is required."
    End If

    ScopeStack_.Add ps
End Sub

' Internal escape hatch for explicit scope pops in specs or advanced code.
'
' Normal callers should not use this directly. `CloseScope` is the usual guard
' path because it sends the final local 100% update before popping.
Friend Sub PopScope(ps As ProgressScope)
    If ScopeStack_.Count > 1 Then
        EnsureTopScope ps, "ProgressManager.PopScope"
        ScopeStack_.Remove ScopeStack_.Count
    End If
End Sub

' Closes the active scope after notifying it at local 100%.
'
' The top-scope check is important: scope guards must be destroyed in LIFO
' order, matching how nested `With Progress.Scope(...)` blocks are written.
Friend Sub CloseScope(ps As ProgressScope)
    EnsureTopScope ps, "ProgressManager.CloseScope"
    Update 1
    ScopeStack_.Remove ScopeStack_.Count
End Sub

' Reports local progress in the active scope.
'
' Notifications are monotonic at the root level. This avoids noisy regressions
' when a child scope starts at local 0 after its parent has already advanced.
Public Sub Update(localProgressValue As Double)
    ValidateProgressValue localProgressValue

    Dim progressValue As Double
    
    progressValue = CalculateProgressValue(localProgressValue)
    If progressValue > LastProgressValue_ Then
        LastProgressValue_ = progressValue
        NotifyListeners progressValue
    End If
End Sub

Private Function CalculateProgressValue(localProgressValue As Double) As Double
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
    
    CalculateProgressValue = offset + ratio * localProgressValue
End Function

Private Sub ValidateProgressValue(progressValue As Double)
    If progressValue < 0 Or progressValue > 1 Then
        Err.Raise 5, "ProgressManager.Update", "Progress value must be between 0 and 1."
    End If
End Sub

Private Sub EnsureRootScope(source As String)
    If ScopeStack_.Count <> 1 Then
        Err.Raise 5, source, "Cannot reset progress while a scope is active."
    End If
End Sub

Private Sub EnsureTopScope(ps As ProgressScope, source As String)
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

Private Sub NotifyListeners(progressValue As Double)
    Dim listener As IProgressListener
    
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
End Sub
