VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ref. <https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raiseevent-statement>
Option Explicit

Private WithEvents ts As TimerState
Attribute ts.VB_VarHelpID = -1
Private Const FinalTime As Double = 10#  ' Second

Private Sub Class_Initialize()
    Set ts = New TimerState
End Sub


Sub Start()
    ts.TimerTask FinalTime
End Sub

Private Sub ts_UpdateElapsedTime(ByVal elapsedTime As Double)
    Debug.Print CStr(Format(elapsedTime, "0.00"))
End Sub

Private Sub ts_DisplayFinalTime()
    Debug.Print "Until now"
    Debug.Print CStr(FinalTime)
End Sub


