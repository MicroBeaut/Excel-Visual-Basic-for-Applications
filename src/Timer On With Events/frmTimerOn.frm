VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimerOn 
   Caption         =   "Timer On"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "frmTimerOn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimerOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents myTimerOn As TimerOn
Attribute myTimerOn.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Set myTimerOn = New TimerOn
End Sub

Private Sub cmdEnable_Click()
    If cmdEnable.Caption = "Enable" Then
        cmdEnable.Caption = "Disable"
        cmdPause.Enabled = True
        cmdCancel.Enabled = True
        myTimerOn.Enable = True
    Else
        cmdEnable.Caption = "Enable"
        cmdPause.Enabled = False
        cmdCancel.Enabled = False
        myTimerOn.Enable = False
        
    End If
End Sub

Private Sub cmdPause_Click()
    myTimerOn.Pause = Not myTimerOn.Pause
    If myTimerOn.Pause Then
        cmdPause.Caption = "Run"
        cmdEnable.Enabled = False
    Else
        cmdPause.Caption = "Pause"
        cmdEnable.Enabled = True
    End If
End Sub

Private Sub CommandButton2_Click()
    myTimerOn.Enable = False
End Sub

Private Sub cmdCancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdEnable.Caption = "Enable"
    cmdEnable.Enabled = True
    cmdPause.Enabled = False
    myTimerOn.Cancellation = True
    myTimerOn.Pause = False
    myTimerOn.Enable = False
End Sub

Private Sub cmdCancel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    myTimerOn.Cancellation = False
End Sub

Sub myTimerOn_OnTimerOn()
    Debug.Print "Event on Timer On"
End Sub

Sub myTimerOn_OnOutputChanged(ByVal value As Boolean)
    If value Then
        lblState.BackColor = vbGreen
        cmdEnable.Enabled = True
        cmdPause.Enabled = False
        cmdCancel.Enabled = False
        myTimerOn.Pause = False
    Else
        lblState.BackColor = vbRed
    End If
End Sub

Sub myTimerOn_OnUpdateElapsedTime(ByVal value As Double)
    lblElapsed.Caption = CStr(Format(value, "0.00"))
End Sub
