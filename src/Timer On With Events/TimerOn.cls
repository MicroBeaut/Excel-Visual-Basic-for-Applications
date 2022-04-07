VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OnUpdateElapsedTime(ByVal value As Double)
Public Event OnTimerOn()
Public Event OnOutputChanged(ByVal vaoue As Boolean)

Private m_Enable As Boolean
Private m_Pause As Boolean
Private m_Cancellation As Boolean
Private m_TimeDelay As Double
Private m_DeltaTime As Double
Private m_TimeInterval As Double

Private m_ElapsedTime As Double
Private m_OutputState As Boolean

Private m_currEnable As Boolean
Private m_prevEnable As Boolean

Private Sub Class_Initialize()
    m_TimeDelay = 3#
    m_TimeInterval = 0.25
End Sub

Property Let Enable(value As Boolean)
    m_prevEnable = m_currEnable
    m_currEnable = value
    m_Enable = value
    If m_currEnable And Not m_prevEnable Then
        TimerTask
    End If
    If Not m_currEnable And m_prevEnable Then
        SetOutput (False)
        m_ElapsedTime = 0#
        RaiseEvent OnUpdateElapsedTime(m_ElapsedTime)
    End If
End Property

Property Get Enable() As Boolean
   Enable = m_Enable = Enable
End Property

Property Let Pause(value As Boolean)
    m_Pause = value
End Property

Property Get Pause() As Boolean
    Pause = m_Pause
End Property

Property Let Cancellation(value As Boolean)
    m_Cancellation = value
End Property

Property Get Cancellation() As Boolean
    Cancellation = m_TimeDelay
End Property

Property Let TimeDelay(value As Double)
    m_TimeDelay = value
End Property

Property Get TimeDelay() As Double
    TimeDelay = m_TimeDelay
End Property

Property Let TimeInterval(value As Double)
    m_TimeInterval = value
End Property

Property Get TimeInterval() As Double
    TimeInterval = m_TimeInterval
End Property

Property Get OutputState() As Boolean
    OutputState = m_OutputState
End Property

Property Get ElapsedTime() As Double
    ElapsedTime = m_ElapsedTime
End Property

Private Function TimerTask()
    Dim currT As Double
    Dim prevT As Double
    Dim currP As Boolean
    Dim prevP As Boolean
    Dim elapsedInterval As Double
    
    If m_currEnable And Not m_prevEnable Then
        m_ElapsedTime = 0
        currT = Timer
    End If
    Do While m_Enable And Not m_Cancellation And Not m_OutputState
        prevP = currP
        currP = m_Pause
        If Not currP And prevP Then
            currT = Timer
        End If
        If Not m_Pause Then
            prevT = currT
            currT = Timer
            m_DeltaTime = currT - prevT
            m_ElapsedTime = m_ElapsedTime + m_DeltaTime
            elapsedInterval = elapsedInterval + m_DeltaTime
            If elapsedInterval > m_TimeInterval Then
                elapsedInterval = 0
                RaiseEvent OnUpdateElapsedTime(m_ElapsedTime)
            End If
        End If
        If m_ElapsedTime > m_TimeDelay Then
            m_ElapsedTime = m_TimeDelay
            SetOutput (True)
            RaiseEvent OnUpdateElapsedTime(m_ElapsedTime)
            RaiseEvent OnTimerOn
        End If
        DoEvents
    Loop
End Function

Private Sub SetOutput(value As Boolean)
    If value <> m_OutputState Then
        m_OutputState = value
        RaiseEvent OnOutputChanged(value)
    End If
End Sub

