VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressScopeGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Lifetime token for a pushed progress scope.
'
' VBA has no `Using` statement for arbitrary objects, but a temporary object in
' a `With ... End With` block stays alive until `End With`. This class uses
' `Class_Terminate` to close the scope when that lifetime ends.
Private Manager_ As ProgressManager
Private Scope_ As ProgressScope
Private Initialized_ As Boolean

' Creates and pushes a scope. Keep this object alive for the desired scope lifetime.
'
' Initialization stores both the manager and the exact scope object so closing
' can verify that this guard still owns the stack top.
Public Sub Init(ByVal manager As ProgressManager, ByVal startPoint As Double, ByVal endPoint As Double)
    If Initialized_ Then
        Err.Raise 5, "ProgressScopeGuard.Init", "Guard is already initialized."
    End If
    If manager Is Nothing Then
        Err.Raise 5, "ProgressScopeGuard.Init", "Progress manager is required."
    End If

    Dim ps As ProgressScope

    Set ps = New ProgressScope
    ps.Init startPoint, endPoint
    
    Set Manager_ = manager
    Set Scope_ = ps
    With Manager_
        .PushScope Scope_
        ' Mark initialized only after the scope has been pushed successfully.
        Initialized_ = True
        .Update 0
    End With
End Sub

Private Sub Class_Initialize()
    Initialized_ = False
End Sub

Private Sub Class_Terminate()
    If Not Initialized_ Then
        Exit Sub
    End If
    If Manager_ Is Nothing Then
        Exit Sub
    End If

    ' Destructors should not raise: termination can happen during error
    ' unwinding, and raising here would obscure the original failure.
    On Error Resume Next
    Manager_.CloseScope Scope_
    If Err.Number <> 0 Then
        Debug.Print "Progress scope close failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub
