VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Public entry point for the library.
'
' This class is intentionally predeclared (`VB_PredeclaredId = True`) so callers
' can write `Progress.Begin`, `Progress.Scope`, and `Progress.Update` without
' importing a standard module or passing a manager through every procedure.
Private Manager_ As ProgressManager

' Exposes the underlying manager for advanced integrations.
'
' Most callers should prefer the facade methods below. The manager is available
' mainly for custom listeners, specs, and host-specific wiring.
Public Property Get Manager() As ProgressManager
    Set Manager = CurrentManager
End Property

' Starts a new root operation and notifies listeners with 0%.
'
' Must be called outside active `Progress.Scope` blocks.
Public Sub Begin()
    CurrentManager.Begin
End Sub

' Finishes the current root operation, notifies listeners with 100%, and resets state.
'
' Must be called outside active `Progress.Scope` blocks. Scope completion is
' handled by `ProgressScopeGuard` when the `With` block ends.
Public Sub Finish()
    CurrentManager.Finish
End Sub

' Resets progress state while keeping registered listeners.
'
' Use this to recover between operations; use `Begin` / `Finish` for the normal
' operation lifecycle.
Public Sub Reset()
    CurrentManager.Reset
End Sub

' Creates a lifetime token that maps local 0..1 progress into startPoint..endPoint.
'
' Intended usage:
'
'   With Progress.Scope(0, 0.3)
'       Progress.Update 0.5
'   End With
'
' The returned guard must remain alive for the scope duration.
Public Function Scope(startPoint As Double, endPoint As Double) As ProgressScopeGuard
    Set Scope = New ProgressScopeGuard
    Scope.Init CurrentManager, startPoint, endPoint
End Function

' Reports local progress in the current scope.
'
' `localProgressValue` must be between 0 and 1. Nested scopes translate this
' local value into the root operation's normalized progress.
Public Sub Update(localProgressValue As Double)
    CurrentManager.Update localProgressValue
End Sub

' Registers or replaces a named listener.
'
' Names make listener setup idempotent: adding the same name again replaces the
' previous listener.
Public Sub AddListener(name As String, listener As IProgressListener)
    CurrentManager.AddListener name, listener
End Sub

' Removes a named listener if it exists.
Public Sub RemoveListener(name As String)
    CurrentManager.RemoveListener name
End Sub

Private Property Get CurrentManager() As ProgressManager
    If Manager_ Is Nothing Then
        Set Manager_ = New ProgressManager
    End If

    Set CurrentManager = Manager_
End Property
