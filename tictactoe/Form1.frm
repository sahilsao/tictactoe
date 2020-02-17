VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tic-Tae-Toe by TRiBe"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   FontTransparent =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   6765
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   4560
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   0
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2295
      ScaleWidth      =   3135
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5475
      TabIndex        =   9
      Top             =   6030
      Width           =   5475
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   3000
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2400
         Top             =   120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1 Score"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer's Score"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.PictureBox c3 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.PictureBox c2 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2160
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox c1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox b3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3960
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox b2 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2280
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox b1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox a3 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox a2 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2160
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox a1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p1score
Dim p2score

Dim player_info
Dim aa1
Dim aa2
Dim aa3
Dim bb1
Dim bb2
Dim bb3
Dim cc1
Dim cc2
Dim cc3

Dim turn

Dim pa1
Dim pa2
Dim pa3
Dim pb1
Dim pb2
Dim pb3
Dim pc1
Dim pc2
Dim pc3



Private Sub Command1_Click()
play "win.wav"


End Sub

Private Sub Command2_Click()
MsgBox pa1 & pa2 & pa3 & vbNewLine & pb1 & pb2 & pb3 & vbNewLine & pc1 & pc2 & pc3

End Sub

Sub checktile(ByVal tile As String, ByVal playercheck As String)

If turn = 0 Then
turn = 1
GoTo 1
End If

If turn = 1 Then
turn = 0
GoTo 1
End If

1 If tile = "a1" Then
If aa1 = 0 Then
If playercheck = 0 Then
colortile "a1", playercheck
pa1 = 0
Else
colortile "a1", playercheck
pa1 = 1
End If
Else
End If
End If

If tile = "a2" Then
If aa2 = 0 Then
If playercheck = 0 Then
pa2 = 0
colortile "a2", playercheck
Else
pa2 = 1
colortile "a2", playercheck
End If
Else
End If
End If

If tile = "a3" Then
If aa3 = 0 Then
If playercheck = 0 Then
pa3 = 0
colortile "a3", playercheck
Else
pa3 = 1
colortile "a3", playercheck
End If
Else
End If
End If


If tile = "b1" Then
If bb1 = 0 Then
If playercheck = 0 Then
pb1 = 0
colortile "b1", playercheck
Else
pb1 = 1
colortile "b1", playercheck
End If
Else
End If
End If


If tile = "b2" Then
If bb2 = 0 Then
If playercheck = 0 Then
pb2 = 0
colortile "b2", playercheck
Else
pb2 = 1
colortile "b2", playercheck
End If
Else
End If
End If


If tile = "b3" Then
If bb3 = 0 Then
If playercheck = 0 Then
pb3 = 0
colortile "b3", playercheck
Else
pb3 = 1
colortile "b3", playercheck
End If
Else
End If
End If


If tile = "c1" Then
If cc1 = 0 Then
If playercheck = 0 Then
pc1 = 0
colortile "c1", playercheck
Else
pc1 = 1
colortile "c1", playercheck
End If
Else
End If
End If


If tile = "c2" Then
If cc2 = 0 Then
If playercheck = 0 Then
pc2 = 0
colortile "c2", playercheck
Else
pc2 = 1
colortile "c2", playercheck
End If
Else
End If
End If


If tile = "c3" Then
If cc3 = 0 Then
If playercheck = 0 Then
pc3 = 0
colortile "c3", playercheck
Else
pc3 = 1
colortile "c3", playercheck
End If
Else
End If
End If


checkwin
If turn = 1 Then
playcomputer
End If



End Sub


Sub colortile(ByVal colortile As String, ByVal playercolor As String)

If colortile = "a1" Then
aa1 = 1
If playercolor = "0" Then
a1.Picture = LoadPicture(App.Path & "/o.jpg")
Else
a1.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If

If colortile = "a2" Then
aa2 = 1
If playercolor = "0" Then
a2.Picture = LoadPicture(App.Path & "/o.jpg")
Else
a2.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If

If colortile = "a3" Then
aa3 = 1
If playercolor = "0" Then
a3.Picture = LoadPicture(App.Path & "/o.jpg")
Else
a3.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If


If colortile = "b1" Then
bb1 = 1
If playercolor = "0" Then
b1.Picture = LoadPicture(App.Path & "/o.jpg")
Else
b1.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If


If colortile = "b2" Then
bb2 = 1
If playercolor = "0" Then
b2.Picture = LoadPicture(App.Path & "/o.jpg")
Else
b2.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If


If colortile = "b3" Then
bb3 = 1
If playercolor = "0" Then
b3.Picture = LoadPicture(App.Path & "/o.jpg")
Else
b3.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If


If colortile = "c1" Then
cc1 = 1
If playercolor = "0" Then
c1.Picture = LoadPicture(App.Path & "/o.jpg")
Else
c1.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If


If colortile = "c2" Then
cc2 = 1
If playercolor = "0" Then
c2.Picture = LoadPicture(App.Path & "/o.jpg")
Else
c2.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If


If colortile = "c3" Then
cc3 = 1
If playercolor = "0" Then
c3.Picture = LoadPicture(App.Path & "/o.jpg")
Else
c3.Picture = LoadPicture(App.Path & "/x.jpg")
End If
End If



End Sub

Private Sub a1_Click()

If pa1 = 1 Then
play "no"
Exit Sub
End If


If turn = 0 Then
checktile "a1", player_info
play "yes"
End If

End Sub

Private Sub a2_Click()
If pa2 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "a2", player_info
play "yes"
End If
End Sub

Private Sub a3_Click()
If pa3 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "a3", player_info
play "yes"
End If
End Sub

Private Sub b1_Click()
If pb1 = 1 Then
play "no"
Exit Sub
End If
If turn = 0 Then
checktile "b1", player_info
play "yes"
End If
End Sub

Private Sub b2_Click()
If pb2 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "b2", player_info
play "yes"
End If
End Sub

Private Sub b3_Click()
If pb3 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "b3", player_info
play "yes"
End If
End Sub

Private Sub c1_Click()
If pc1 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "c1", player_info
play "yes"
End If
End Sub

Private Sub c2_Click()
If pc2 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "c2", player_info
play "yes"
End If
End Sub

Private Sub c3_Click()
If pc3 = 1 Then
play "no"
Exit Sub
End If

If turn = 0 Then
checktile "c3", player_info
play "yes"
End If
End Sub

Private Sub Form_Loadv()
player_info = 1
cleargame
End Sub

Private Sub Text1_Change()
player_info = Text1.Text
End Sub



Sub checkwin()

'CHECK BLUE
'000
'---
'---
If pa1 = 1 Then
    If pa2 = 1 Then
        If pa3 = 1 Then
        player1win
        End If
    End If
End If

'CHECK BLUE
'---
'000
'---
If pb1 = 1 Then
    If pb2 = 1 Then
        If pb3 = 1 Then
        player1win
        End If
    End If
End If



'CHECK BLUE
'---
'---
'000
If pc1 = 1 Then
    If pc2 = 1 Then
        If pc3 = 1 Then
        player1win
        End If
    End If
End If


'CHECK BLUE
'0--
'0--
'0--
If pa1 = 1 Then
    If pb1 = 1 Then
        If pc1 = 1 Then
        player1win
        End If
    End If
End If

'CHECK BLUE
'-0-
'-0-
'-0-
If pa2 = 1 Then
    If pb2 = 1 Then
        If pc2 = 1 Then
        player1win
        End If
    End If
End If


'CHECK BLUE
'--0
'--0
'--0
If pa3 = 1 Then
    If pb3 = 1 Then
        If pc3 = 1 Then
        player1win
        End If
    End If
End If

'CHECK BLUE
'0--
'-0-
'--0
If pa1 = 1 Then
    If pb2 = 1 Then
        If pc3 = 1 Then
        player1win
        End If
    End If
End If

'CHECK BLUE
'--0
'-0-
'0--
If pa3 = 1 Then
    If pb2 = 1 Then
        If pc1 = 1 Then
        player1win
        End If
    End If
End If

'--------------------------------------------------------------------------------

'CHECK RED
'000
'---
'---
If pa1 = 0 Then
    If pa2 = 0 Then
        If pa3 = 0 Then
        player2win 1
        End If
    End If
End If

'CHECK RED
'---
'000
'---
If pb1 = 0 Then
    If pb2 = 0 Then
        If pb3 = 0 Then
        player2win 2
        End If
    End If
End If



'CHECK RED
'---
'---
'000
If pc1 = 0 Then
    If pc2 = 0 Then
        If pc3 = 0 Then
        player2win 3
        End If
    End If
End If


'CHECK RED
'0--
'0--
'0--
If pa1 = 0 Then
    If pb1 = 0 Then
        If pc1 = 0 Then
        player2win 4
        End If
    End If
End If

'CHECK RED
'-0-
'-0-
'-0-
If pa2 = 0 Then
    If pb2 = 0 Then
        If pc2 = 0 Then
        player2win 5
        End If
    End If
End If


'CHECK RED
'--0
'--0
'--0
If pa3 = 0 Then
    If pb3 = 0 Then
        If pc3 = 0 Then
        player2win 6
        End If
    End If
End If

'CHECK RED
'0--
'-0-
'--0
If pa1 = 0 Then
    If pb2 = 0 Then
        If pc3 = 0 Then
       player2win 7
        End If
    End If
End If

'CHECK RED
'--0
'-0-
'0--
If pa3 = 0 Then
    If pb2 = 0 Then
        If pc1 = 0 Then
        player2win 8
        End If
    End If
End If

End Sub


Sub player1win()
play "win"
playerdisabled
cleargame
p1score = p1score + 100
p2score = p2score - 50
Label3.Caption = p1score
Label4.Caption = p2score
cmsg "Good JOB!"
'Timer1.Enabled = True

End Sub
Sub player2win(ByVal az)
play "lost"
playerdisabled
cleargame
p2score = p2score + 100
p1score = p1score - 50
Label3.Caption = p1score
Label4.Caption = p2score
cmsg "Sorry, you lost!"
'Timer1.Enabled = True

End Sub


Sub cleargame()

a1.Picture = LoadPicture(App.Path & "/blank.jpg")
a2.Picture = LoadPicture(App.Path & "/blank.jpg")
a3.Picture = LoadPicture(App.Path & "/blank.jpg")
b1.Picture = LoadPicture(App.Path & "/blank.jpg")
b2.Picture = LoadPicture(App.Path & "/blank.jpg")
b3.Picture = LoadPicture(App.Path & "/blank.jpg")
c1.Picture = LoadPicture(App.Path & "/blank.jpg")
c2.Picture = LoadPicture(App.Path & "/blank.jpg")
c3.Picture = LoadPicture(App.Path & "/blank.jpg")


aa1 = 0
aa2 = 0
aa3 = 0
bb1 = 0
bb2 = 0
bb3 = 0
cc1 = 0
cc2 = 0
cc3 = 0

pa1 = 3
pa2 = 3
pa3 = 3
pb1 = 3
pb2 = 3
pb3 = 3
pc1 = 3
pc2 = 3
pc3 = 3

End Sub

Private Sub Form_Load()
p1score = 0
p2score = 0
Label3.Caption = p1score
Label4.Caption = p2score

turn = 0
player_info = 1
cleargame
End Sub


Sub playcomputer()

a1.Enabled = False
a2.Enabled = False
a3.Enabled = False
b1.Enabled = False
b2.Enabled = False
b3.Enabled = False
c1.Enabled = False
c2.Enabled = False
c3.Enabled = False


Do

If aa1 = 1 Then
If aa2 = 1 Then
If aa3 = 1 Then
If bb1 = 1 Then
If bb2 = 1 Then
If bb3 = 1 Then
If cc1 = 1 Then
If cc2 = 1 Then
If cc3 = 1 Then
'draw
cmsg "DRAW!"
cleargame


End If
End If
End If
End If
End If
End If
End If
End If
End If


Randomize Timer
df = Int(Rnd * 9) + 1






If df = 1 Then
If aa1 = 0 Then
checktile "a1", 0
enabledplayer
Exit Sub
End If
End If

If df = 2 Then
If aa2 = 0 Then
checktile "a2", 0
enabledplayer
Exit Sub
End If
End If

If df = 3 Then
If aa3 = 0 Then
checktile "a3", 0
enabledplayer
Exit Sub
End If
End If

If df = 4 Then
If bb1 = 0 Then
checktile "b1", 0
enabledplayer
Exit Sub
End If
End If

If df = 5 Then
If bb2 = 0 Then
checktile "b2", 0
enabledplayer
Exit Sub
End If
End If

If df = 6 Then
If bb3 = 0 Then
checktile "b3", 0
enabledplayer
Exit Sub
End If
End If

If df = 7 Then
If cc1 = 0 Then
checktile "c1", 0
enabledplayer
Exit Sub
End If
End If

If df = 8 Then
If cc2 = 0 Then
checktile "c2", 0
enabledplayer
Exit Sub
End If
End If

If df = 9 Then
If cc3 = 0 Then
checktile "c3", 0
enabledplayer
Exit Sub
End If
End If




Loop

End Sub
Sub enabledplayer()
turn = 0

a1.Enabled = True
a2.Enabled = True
a3.Enabled = True
b1.Enabled = True
b2.Enabled = True
b3.Enabled = True
c1.Enabled = True
c2.Enabled = True
c3.Enabled = True
End Sub

Private Sub Timer1_Timer()
cleargame
Timer1.Enabled = False

End Sub
Sub playerdisabled()

a1.Enabled = False
a2.Enabled = False
a3.Enabled = False
b1.Enabled = False
b2.Enabled = False
b3.Enabled = False
c1.Enabled = False
c2.Enabled = False
c3.Enabled = False


End Sub


Sub cmsg(ByVal mor As String)

Picture2.Visible = True
Label5.Caption = mor
Timer2.Enabled = True



End Sub

Private Sub Timer2_Timer()
Picture2.Visible = False
Timer2.Enabled = False

End Sub
Sub play(ByVal filez As String)
MMControl1.Command = "close"
MMControl1.Command = "stop"
MMControl1.Command = "prev"
MMControl1.FileName = App.Path & "\" & filez & ".wav"
MMControl1.Command = "open"
MMControl1.Command = "play"
End Sub
