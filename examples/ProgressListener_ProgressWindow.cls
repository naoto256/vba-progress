VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressListener_ProgressWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IProgressListener

' Optional adapter for projects that provide a `ProgressWindow` form.
'
' The core library does not require this file. Import it only when the host
' project defines a compatible form with `OpenUI`, `SetProgressValue`, and
' `CloseUI` members.
#If VBA7 Then
    #If Mac Then
        Private Declare PtrSafe Function usleep Lib "libc.dylib" (ByVal microseconds As Long) As Long
    #Else
        Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    #End If
#Else
    #If Mac Then
        Private Declare Function usleep Lib "libc.dylib" (ByVal microseconds As Long) As Long
    #Else
        Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    #End If
#End If

Private IsProgressBarOpen_ As Boolean

Private Sub Class_Initialize()
    IsProgressBarOpen_ = False
End Sub

Private Sub IProgressListener_OnProgressUpdated(progressValue As Double)
    With ProgressWindow
        If Not IsProgressBarOpen_ Then
            .OpenUI
            IsProgressBarOpen_ = True
        End If
        
        .SetProgressValue progressValue
        
        If progressValue = 1 Then
            SleepMilliseconds 500
            .CloseUI
            IsProgressBarOpen_ = False
        End If
    End With
End Sub

Private Sub SleepMilliseconds(milliseconds As Long)
    If milliseconds <= 0 Then
        Exit Sub
    End If

    ' Use the host OS sleep primitive. A timer loop would waste CPU and can make
    ' Excel feel worse exactly when the UI is trying to settle after completion.
#If Mac Then
    usleep milliseconds * 1000&
#Else
    Sleep milliseconds
#End If
End Sub
