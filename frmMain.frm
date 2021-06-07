VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O's and X's by Paul Blower"
   ClientHeight    =   4545
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ResetTimer 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4560
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   4230
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6030
      LocalPort       =   6030
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Speak"
      Default         =   -1  'True
      Height          =   285
      Left            =   6900
      TabIndex        =   3
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox SpBox 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   4200
      Width           =   3735
   End
   Begin RichTextLib.RichTextBox StatusBox 
      Height          =   1410
      Left            =   60
      TabIndex        =   1
      Top             =   3090
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   2487
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin RichTextLib.RichTextBox ChatBox 
      Height          =   4065
      Left            =   3105
      TabIndex        =   0
      Top             =   60
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   7170
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmMain.frx":00D5
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   8
      Left            =   2055
      Picture         =   "frmMain.frx":01AA
      Top             =   2055
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   7
      Left            =   1095
      Picture         =   "frmMain.frx":0469
      Top             =   2055
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   6
      Left            =   135
      Picture         =   "frmMain.frx":0728
      Top             =   2055
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   5
      Left            =   2055
      Picture         =   "frmMain.frx":09E7
      Top             =   1095
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   4
      Left            =   1095
      Picture         =   "frmMain.frx":0CA6
      Top             =   1095
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   3
      Left            =   135
      Picture         =   "frmMain.frx":0F65
      Top             =   1095
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   2
      Left            =   2055
      Picture         =   "frmMain.frx":1224
      Top             =   135
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   1
      Left            =   1095
      Picture         =   "frmMain.frx":14E3
      Top             =   135
      Width           =   900
   End
   Begin VB.Image XDisp 
      Height          =   900
      Index           =   0
      Left            =   135
      Picture         =   "frmMain.frx":17A2
      Top             =   135
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   75
      Top             =   75
      Width           =   2940
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect Screen"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Disconnect"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'           /-----------------------------------\
'           |  O's and X's by Paul Blower       |
'           |                                   |
'           | Hi, i hope you like this game..   |
'           | the codes probably a bit messy    |
'           | but i think its understandable    |
'           | ive tried to comment everything.  |
'           | its my first submission, so       |
'           | PLEASE take the time to vote for  |
'           | this if you like It!              |
'           \-----------------------------------/

Dim OXArray(8)
Dim HasWon As Boolean

Dim MPReady As Boolean
Dim MPMyGo As Boolean
Dim MPMySymbol As String
Dim MPOtherSymbol As String
Dim MPMyScore
Dim MPOtherScore
Public MPMyName
Dim MPOtherName

Private Sub Command1_Click() 'just code for chat
ChatBox.Text = ChatBox.Text & MPMyName & ": " & SpBox.Text & vbCrLf
ws.SendData "OXCHAT" & SpBox.Text
ChatBox.SelStart = Len(ChatBox.Text)
SpBox.Text = ""
SpBox.SetFocus
End Sub

Private Sub SpBox_KeyDown(KeyCode As Integer, Shift As Integer) 'if you press enter, send chat
If KeyCode = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Form_Load()
MPMyScore = 0
MPOtherScore = 0
HasWon = False
MPReady = False
End Sub

Sub StartGame() 'This is bad coding i know, but i really cant be bothered to sort it out :p
MPReady = True  'knowing my luck, id mess up the game
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub mnuClose_Click() 'Disconnects us
ws.Close
End Sub

Private Sub mnuConnect_Click() 'Just shows form2 (the connection form)
Load Form2
Form2.Show
End Sub

Private Sub ResetTimer_Timer() 'This sub resets the whole game

For i = 0 To 8 'This loop makes the drawing area blank
OXArray(i) = 0
XDisp(i).Picture = LoadPicture(App.Path & "\blank.jpg")
Next

ResetTimer.Enabled = False 'basically, the code below just sets everything back to how
If MPMyGo Then             'it was at the start of the game. Also, it makes the loser
MPMyGo = True              'have 1st go :)
MPReady = True
HasWon = False
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Game Reset." & vbCrLf
StatusBox.Text = StatusBox.Text & "Your Go" & vbCrLf

Else
MPMyGo = False
MPReady = True
HasWon = False
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Game Reset." & vbCrLf
StatusBox.Text = StatusBox.Text & "Other Sides Go" & vbCrLf
End If
End Sub

Private Sub ws_Close()
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Disconnected." 'Tell us if were disconnected and stop play
MPReady = False
End Sub

Private Sub ws_Connect()
MPOtherSymbol = "x" 'Makes you O's and the other side X's
MPMySymbol = "o"

MPMyGo = False
ws.SendData "OXNAME" & MPMyName 'tell the other side your name
App.Title = App.Title & " - Playing As " & MPMyName 'Just looks more professional :)

StatusBox.Text = StatusBox.Text & "Connected to " & ws.RemoteHostIP & vbCrLf
StartGame
StatusBox.Text = StatusBox.Text & "Game Started. You are " & UCase(MPMySymbol) & "'s" & vbCrLf
StatusBox.Text = StatusBox.Text & "Other Sides Go" & vbCrLf

End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
MPOtherSymbol = "o" 'Makes you X's and the other side O's
MPMySymbol = "x"
ws.Close
ws.Accept requestID 'Accept the conection

MPMyGo = True
ws.SendData "OXNAME" & MPMyName 'tell the other side your name
App.Title = App.Title & " - Playing As " & MPMyName 'Just looks more professional :)

StatusBox.Text = StatusBox.Text & "Connected to " & ws.RemoteHostIP & vbCrLf
StartGame
StatusBox.Text = StatusBox.Text & "Game Started. You are " & UCase(MPMySymbol) & "'s" & vbCrLf
StatusBox.Text = StatusBox.Text & "Your Go" & vbCrLf

End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim Wdata As String
ws.GetData Wdata

Dim IndexClick As Integer

If InStr(1, Wdata, "OXDRAW") Then 'If its a draw then this gets sent

    IndexClick = Int(Mid(Wdata, 8, 1)) 'Draw their move
    OXArray(IndexClick) = MPOtherSymbol
    XDisp(IndexClick).Picture = LoadPicture(App.Path & "\" & MPOtherSymbol & ".jpg")

StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Its A Draw!!" & vbCrLf
StatusBox.Text = StatusBox.Text & "Scores are, You: " & MPMyScore & ", " & MPOtherName & ":" & MPOtherScore & vbCrLf

HasWon = False
MPReady = False
MPMyGo = False
ResetTimer.Enabled = True 'reset the game

Exit Sub
End If

If InStr(1, Wdata, "OXWIN") Then 'if the other side wins, this gets sent
HasWon = True
MPReady = False
MPMyGo = True
MPOtherScore = MPOtherScore + 1 'Add onto their score

    IndexClick = Int(Mid(Wdata, 8, 1)) 'Draw their move
    OXArray(IndexClick) = MPOtherSymbol
    XDisp(IndexClick).Picture = LoadPicture(App.Path & "\" & MPOtherSymbol & ".jpg")

StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Other Side Wins!!" & vbCrLf
StatusBox.Text = StatusBox.Text & "Scores are, You: " & MPMyScore & ", " & MPOtherName & ":" & MPOtherScore & vbCrLf
ResetTimer.Enabled = True 'reset the game
Exit Sub
End If

If Left(Wdata, 7) = "OXCLICK" Then 'Draw the other players move, and make it your go
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Receiving Move..." & vbCrLf
StatusBox.Text = StatusBox.Text & "Your Go" & vbCrLf
IndexClick = Int(Right(Wdata, Len(Wdata) - 7))
OXArray(IndexClick) = MPOtherSymbol
XDisp(IndexClick).Picture = LoadPicture(App.Path & "\" & MPOtherSymbol & ".jpg")
MPMyGo = True
End If

If Left(Wdata, 6) = "OXNAME" Then 'This is how we know the other players name.
MPOtherName = Right(Wdata, Len(Wdata) - 6)
ChatBox.Text = ChatBox.Text & MPOtherName & " Enters" & vbCrLf
End If

If Left(Wdata, 6) = "OXCHAT" Then 'Incoming data is chat
ChatBox.Text = ChatBox.Text & MPOtherName & ": " & Right(Wdata, Len(Wdata) - 6) & vbCrLf
ChatBox.SelStart = Len(ChatBox.Text)
End If

End Sub

Private Sub XDisp_Click(Index As Integer)
If Not MPReady Or Not MPMyGo Then Exit Sub  'Only let you click on the game area if the game is ready
If Not HasWon Then 'To make sure people cant play after winning!

If OXArray(Index) = 0 Then
OXArray(Index) = MPMySymbol
XDisp(Index).Picture = LoadPicture(App.Path & "\" & MPMySymbol & ".jpg") 'Display a o or an x

MPMyGo = False 'Youve had your go now!

Dim XS As String ' This loop creates a string to check if a player has won
Dim OS As String
For i = 0 To 8
If OXArray(i) = "x" Then
XS = XS & i
ElseIf OXArray(i) = "o" Then
OS = OS & i
End If
Next

'Now we send our move to the other player... OXCLICK is just a protocol followed by where we clicked
ws.SendData "OXCLICK" & Index
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Sending Move..." & vbCrLf
StatusBox.Text = StatusBox.Text & "Other Sides Go" & vbCrLf

'The bit below checks if either side has won, i think its the only way to do it!
'its not as bad as it looks, very simple really

'012    <-- | Thats how the board is laid out. the code below just checks if there are X's on
'345        | 0,1 and 2 or 0,4 and 8 (which would mean theyd won) or o's on the relevant
'678        | places, and if there are, it calls the <player>WIN sub

If InStr(1, XS, "0") > 0 And InStr(1, XS, "1") > 0 And InStr(1, XS, "2") > 0 Then Call XWIN
If InStr(1, XS, "3") > 0 And InStr(1, XS, "4") > 0 And InStr(1, XS, "5") > 0 Then Call XWIN
If InStr(1, XS, "6") > 0 And InStr(1, XS, "7") > 0 And InStr(1, XS, "8") > 0 Then Call XWIN
If InStr(1, XS, "0") > 0 And InStr(1, XS, "3") > 0 And InStr(1, XS, "6") > 0 Then Call XWIN
If InStr(1, XS, "1") > 0 And InStr(1, XS, "4") > 0 And InStr(1, XS, "7") > 0 Then Call XWIN
If InStr(1, XS, "2") > 0 And InStr(1, XS, "5") > 0 And InStr(1, XS, "8") > 0 Then Call XWIN
If InStr(1, XS, "0") > 0 And InStr(1, XS, "4") > 0 And InStr(1, XS, "8") > 0 Then Call XWIN
If InStr(1, XS, "2") > 0 And InStr(1, XS, "4") > 0 And InStr(1, XS, "6") > 0 Then Call XWIN

If InStr(1, OS, "0") > 0 And InStr(1, OS, "1") > 0 And InStr(1, OS, "2") > 0 Then Call OWIN
If InStr(1, OS, "3") > 0 And InStr(1, OS, "4") > 0 And InStr(1, OS, "5") > 0 Then Call OWIN
If InStr(1, OS, "6") > 0 And InStr(1, OS, "7") > 0 And InStr(1, OS, "8") > 0 Then Call OWIN
If InStr(1, OS, "0") > 0 And InStr(1, OS, "3") > 0 And InStr(1, OS, "6") > 0 Then Call OWIN
If InStr(1, OS, "1") > 0 And InStr(1, OS, "4") > 0 And InStr(1, OS, "7") > 0 Then Call OWIN
If InStr(1, OS, "2") > 0 And InStr(1, OS, "5") > 0 And InStr(1, OS, "8") > 0 Then Call OWIN
If InStr(1, OS, "0") > 0 And InStr(1, OS, "4") > 0 And InStr(1, OS, "8") > 0 Then Call OWIN
If InStr(1, OS, "2") > 0 And InStr(1, OS, "4") > 0 And InStr(1, OS, "6") > 0 Then Call OWIN

'Now we check if its a draw by checing how many o's ans x' are on the board...
If (Len(OS) + Len(XS)) = 9 Then Call ADraw

End If
End If
End Sub

Sub ADraw() 'Gets called when its a draw
ws.SendData "OXDRAW" 'Send Draw protocol
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "Its A Draw!!" & vbCrLf
StatusBox.Text = StatusBox.Text & "Scores are, You: " & MPMyScore & ", " & MPOtherName & ":" & MPOtherScore & vbCrLf
MPReady = False
HasWon = True
MPMyGo = True
ResetTimer.Enabled = True 'Reset the game
End Sub

Sub XWIN()
If Not HasWon Then 'just to make sure won isnt called twice when you double win (the middle!)
HasWon = True

ws.SendData "OXWIN" 'Sends a protocol saying i've won
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "You Win!!" & vbCrLf
MPMyScore = MPMyScore + 1 'Add to your score, then display scores
StatusBox.Text = StatusBox.Text & "Scores are, You: " & MPMyScore & ", " & MPOtherName & ":" & MPOtherScore & vbCrLf

ResetTimer.Enabled = True 'reset the game
End If
End Sub

Sub OWIN()
If Not HasWon Then 'just to make sure won isnt called twice when you double win (the middle!)
HasWon = True

ws.SendData "OXWIN" 'Sends a protocol saying i've won
StatusBox.Text = ""
StatusBox.Text = StatusBox.Text & "You Win!!" & vbCrLf
MPMyScore = MPMyScore + 1 'Add to your score, then display scores
StatusBox.Text = StatusBox.Text & "Scores are, You: " & MPMyScore & ", " & MPOtherName & ":" & MPOtherScore & vbCrLf

ResetTimer.Enabled = True 'reset the game
End If
End Sub
