VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PROJECTILES"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8640
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9240
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "10.0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Text            =   "5"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Height"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "S=4.9*t^2"
      Height          =   255
      Left            =   13680
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   " Horizontal Speed (U)="
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   1440
      X2              =   1320
      Y1              =   240
      Y2              =   120
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   1320
      X2              =   1440
      Y1              =   360
      Y2              =   240
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   840
      X2              =   1440
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Ball 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   3  'Circle
      Top             =   660
      Width           =   495
   End
   Begin VB.Line vv 
      X1              =   960
      X2              =   960
      Y1              =   1160
      Y2              =   11160
   End
   Begin VB.Line hh 
      X1              =   0
      X2              =   2040
      Y1              =   1160
      Y2              =   1160
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BallTop As Long, StartTime As Long, VDirection As String, speed As Long

Private Sub Command1_Click()
'Text2 = (11160 - 1160) / 1000
vv.Y1 = -(1000 * Text2 - 11160)
hh.Y1 = vv.Y1
hh.Y2 = vv.Y1
Ball.Left = 0
Ball.Top = vv.Y1 - Ball.Height
'MsgBox vv.Y1
End Sub

Private Sub Command2_Click()
Command1_Click
Timer1 = False
StartTime = 0
BallTop = 1
Timer1 = True
Ball.Left = 0
Timer2 = True
End Sub

Private Sub Timer1_Timer()

If StartTime = 0 Then
StartTime = Timer
BallTop = Ball.Top
VDirection = "down"
speed = 0
End If

If VDirection = "down" Then
DoEvents
speed = speed + 9.8 * (Timer - StartTime)
Ball.Top = BallTop + (4.9 * (Timer - StartTime) ^ 2) * 700

If Ball.Top >= Me.Height - Ball.Width Then
Ball.Top = Me.Height - Ball.Height * 2
VDirection = "up"
speed = speed - Text3
BallTop = Ball.Top

Timer1 = False
Timer2 = False
Me.Caption = "Time Spent=" & Round((Timer - StartTime), 3)
StartTime = Timer
End If

End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Ball.Left = (Timer - StartTime) * Text1 * 700
End Sub
