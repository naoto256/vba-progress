VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Listener interface for progress sinks.
'
' Implement this in UI, logging, or test adapters. Values are normalized root
' progress values from 0 to 1.
Public Sub OnProgressUpdated(ByVal progressValue As Double)
End Sub
