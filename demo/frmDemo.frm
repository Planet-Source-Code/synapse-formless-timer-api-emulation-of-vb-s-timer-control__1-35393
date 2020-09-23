VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer Demo"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Disable Timer"
      Height          =   330
      Left            =   2430
      TabIndex        =   2
      Tag             =   "DISABLE"
      Top             =   165
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Text            =   "1000"
      Top             =   165
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Interval"
      Height          =   330
      Left            =   1065
      TabIndex        =   0
      Top             =   150
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Number of ticks: 0"
      Height          =   270
      Left            =   195
      TabIndex        =   3
      Top             =   660
      Width           =   3630
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objTimer As clsTimer
Attribute objTimer.VB_VarHelpID = -1

Private Sub Command1_Click()
    If Command1.Tag = "DISABLE" Then
        objTimer.Enabled = False
        Command1.Caption = "Enable Timer"
        Command1.Tag = "ENABLE"
    Else
        objTimer.Enabled = True
        Command1.Caption = "Disable Timer"
        Command1.Tag = "DISABLE"
    End If
End Sub

Private Sub Command2_Click()
    objTimer.Interval = CInt(Text1)
End Sub

Private Sub Form_Load()
    Set objTimer = New clsTimer
    objTimer.Interval = 1000 '1 second delay
End Sub

Private Sub objTimer_Timer()
   Static lngX As Long
   
   lngX = lngX + 1
   Label1.Caption = "Number of ticks: " & lngX
End Sub
