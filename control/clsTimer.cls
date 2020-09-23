VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'/*******************************************************
'*  Formless Timer: API emulation of VB Timer control   *
'*  Written by Jahufar Sadique [AKA sYNAPSE]            *
'*  Last updated on May 29, 2002                        *
'*******************************************************/

'/*************************************************
'* Revision History:                              *
'*   1.0.0 [29/05/2002]: Initial build.           *
'*************************************************/

Option Explicit

Public Event Timer()
Attribute Timer.VB_Description = "Occurs when a preset interval for the Timer object has elapsed."

Private intTimerID As Integer
Private intInterval As Integer
Private boolEnabled As Boolean
Private boolTicking  As Boolean

Private WithEvents objTimerEventSink As clsTimerEventSink
Attribute objTimerEventSink.VB_VarHelpID = -1

Private Function StartTimer(Interval As Integer) As Boolean
    intInterval = Interval
    If boolTicking Then ' kill the timer first.
        KillTimer 0, intTimerID
    End If
    
    boolTicking = True
    intTimerID = SetTimer(0, 0, intInterval, AddressOf TimerProc)
    StartTimer = intTimerID
End Function

Private Function StopTimer(TimerID As Integer) As Boolean
    If boolTicking Then _
        StopTimer = KillTimer(0, TimerID) _
    Else _
        StopTimer = True
End Function

Public Property Let Interval(propInterval As Integer)
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    intInterval = propInterval 'remember new interval
    
    If Not (boolEnabled) Then Exit Property 'already disabled.
    
    If intInterval = 0 Then 'stop the timer.
        boolTicking = False
        If intTimerID > 0 Then KillTimer 0, intTimerID
        Exit Property
    End If
    
    StartTimer intInterval
End Property

Public Property Get Interval() As Integer
    Interval = intInterval
End Property

Private Property Get TimerID() As Integer
    TimerID = intTimerID
End Property

Public Property Let Enabled(propEnabled As Boolean)
Attribute Enabled.VB_Description = " Returns/sets a value that determines whether an object can respond to user-generated events."
    'make a copy of the enabled status..
    boolEnabled = propEnabled
    If intInterval <= 0 Then Exit Property
    
    If Not (boolEnabled) Then 'disable (stop) the timer.
        If Not (boolTicking) Then Exit Property 'already disabled!
        'stop the timer.
        boolTicking = False
        KillTimer 0, intTimerID
        Exit Property
    End If
    
    'enable (start) the timer.
    StartTimer intInterval
End Property

Public Property Get Enabled() As Boolean
    Enabled = boolEnabled
End Property

Private Sub objTimerEventSink_Timer()
    RaiseEvent Timer
End Sub

Private Sub Class_Initialize()
    Set objTimerEventSink = TimerEventSink 'create the EventSink
    boolEnabled = True 'the timer is enabled (though it's interval is initially 0)
End Sub

Private Sub Class_Terminate()
    'make sure to stop the timer if the timer exists..
    If (boolTicking) Then StopTimer intTimerID
    Set objTimerEventSink = Nothing
End Sub

