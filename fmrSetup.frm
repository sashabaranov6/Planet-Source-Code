VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Setup"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Server"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Anonymous"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your name below, and then choose to be the server, or to connect."
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   4920
      X2              =   2400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   1440
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form1
Form1.ws.LocalPort = 5050
Form1.ws.Listen
Form1.StatusBox.Text = Form1.StatusBox.Text & "Awaiting Connection..." & vbCrLf
Form1.MPMyName = Text1.Text
Form1.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Load Form1
Form1.ws.RemoteHost = Text2.Text
Form1.ws.RemotePort = 5050
Form1.StatusBox.Text = Form1.StatusBox.Text & "Conecting to " & Text2.Text & "..." & vbCrLf
Form1.ws.Connect
Form1.MPMyName = Text1.Text
Form1.Show
Me.Hide
End Sub

Private Sub Form_Load()
Text2.Text = Form1.ws.LocalIP
End Sub


