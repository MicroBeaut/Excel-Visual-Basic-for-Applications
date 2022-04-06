VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ref. <https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raiseevent-statement>

Option Explicit

Public Event UpdateElapsedTime(ByVal elapsedTime As Double)
Public Event DisplayFinalTime()
Private Const delta As Double = 0.01

Private Const IntervalUpdate As Double = 0.25
Private elapsedInterval As Double

Public Sub TimerTask(ByVal Duration As Double)
    Dim startTime As Double
    startTime = Timer
    Dim timeElapsedSoFar As Double
    timeElapsedSoFar = startTime

    Do While Timer < startTime + Duration
        If Timer - timeElapsedSoFar >= delta Then
            timeElapsedSoFar = timeElapsedSoFar + delta
            elapsedInterval = elapsedInterval + delta
            If (elapsedInterval >= IntervalUpdate) Then
                elapsedInterval = 0
                RaiseEvent UpdateElapsedTime(Timer - startTime)
            End If
            DoEvents
        End If
    Loop
    RaiseEvent DisplayFinalTime
End Sub
