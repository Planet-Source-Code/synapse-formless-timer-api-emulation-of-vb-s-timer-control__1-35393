Attribute VB_Name = "modTimer"
Option Explicit

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private objTimerEventSink As clsTimerEventSink

Public Property Get TimerEventSink() As clsTimerEventSink
    'initialize the EventSink..
    If objTimerEventSink Is Nothing Then _
        Set objTimerEventSink = New clsTimerEventSink
        
    Set TimerEventSink = objTimerEventSink
End Property

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nID As Long, ByVal dwTime As Long)
    objTimerEventSink.Tick
End Sub
