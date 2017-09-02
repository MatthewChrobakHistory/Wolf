VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Wolf"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstIndex 
      Height          =   7665
      Left            =   0
      TabIndex        =   42
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame fraAdminPanel 
      Caption         =   "Admin Panel"
      Height          =   1935
      Left            =   7800
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CheckBox chkACupid 
         Caption         =   "Cupid set lovers?"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdAdmin 
      Caption         =   "Admin Login"
      Height          =   375
      Left            =   5160
      TabIndex        =   39
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General Information"
      Height          =   1815
      Left            =   2880
      TabIndex        =   31
      Top             =   120
      Width           =   4335
      Begin VB.Label lblPlayers 
         Caption         =   "Players: Null"
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblRole 
         Caption         =   "Role's Turn: Null"
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblTime 
         Caption         =   "Time: Day"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Next Turn"
      Height          =   375
      Left            =   2880
      TabIndex        =   30
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame fraPlayer4 
      Caption         =   "Player 4"
      Height          =   1815
      Left            =   7320
      TabIndex        =   26
      Top             =   2040
      Width           =   5655
      Begin VB.Label lblInLove4 
         Caption         =   "In Love: No"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPlayerAlive4 
         Caption         =   "Alive: "
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label lblPlayerRole4 
         Caption         =   "Role: "
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label lblPlayerName4 
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame fraPlayer3 
      Caption         =   "Player 3"
      Height          =   1815
      Left            =   7320
      TabIndex        =   22
      Top             =   2040
      Width           =   5655
      Begin VB.Label lblInLove3 
         Caption         =   "In Love: No"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPlayerName3 
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label lblPlayerRole3 
         Caption         =   "Role: "
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label LblPlayerAlive3 
         Caption         =   "Alive: "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   5055
      End
   End
   Begin VB.Frame fraPlayer2 
      Caption         =   "Player 2"
      Height          =   1815
      Left            =   7320
      TabIndex        =   18
      Top             =   2040
      Width           =   5655
      Begin VB.Label lblInLove2 
         Caption         =   "In Love: No"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPlayerName2 
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LblPlayerRole2 
         Caption         =   "Role: "
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblPlayerAlive2 
         Caption         =   "Alive: "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame frmPlayerList 
      Height          =   1815
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdPlayer12 
         Caption         =   "Command2"
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer11 
         Caption         =   "Command2"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer10 
         Caption         =   "Command2"
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer9 
         Caption         =   "Command2"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer8 
         Caption         =   "Command2"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer4 
         Caption         =   "Player4"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer7 
         Caption         =   "Command2"
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer3 
         Caption         =   "Player3"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer6 
         Caption         =   "Command2"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer2 
         Caption         =   "Player2"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer5 
         Caption         =   "Command2"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayer1 
         Caption         =   "Player1"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame fraPlayer1 
      Caption         =   "Player 1"
      Height          =   1815
      Left            =   7320
      TabIndex        =   1
      Top             =   2040
      Width           =   5655
      Begin VB.Label lblInLove1 
         Caption         =   "In Love: No"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPlayerAlive1 
         Caption         =   "Alive: "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblPlayerRole1 
         Caption         =   "Role: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblPlayerName1 
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmin_Click()

frmAdminLogin.Show

End Sub

Private Sub cmdPlay_Click()
Dim Role As String
Dim msg, sapi
Set sapi = CreateObject("sapi.spvoice")
Dim Name As String
'CT stands for current turn

'exit if a turn is in progress
If cmdPlay.Caption = "Turn in progress" Then Exit Sub


    '//////////////////////////
    '/////CT Turn: Cupid///////
    '//////////////////////////

    If lblRole.Caption = "Role's Turn: Null" Then
        'make sure cupid didn't go yet
        If chkACupid.Value = 0 Then
            If lblTime.Caption = "Time: Day" Then lblTime.Caption = "Time: Night"
            lstIndex.AddItem "Turns to night"
            Role = "Role: Cupid"
            
            'speaking stuff
            msg = "It is now night time."
            sapi.speak msg
            msg = "I call upon cupid to wake up. Will cupid please choose two people to fall in love."
            sapi.speak msg
            
            'look for the cupid and make it his turn
            If lblPlayerRole1.Caption = Role Then
                Call CupidsTurn
            End If
            If LblPlayerRole2.Caption = Role Then
                Call CupidsTurn
            End If
            If lblPlayerRole3.Caption = Role Then
                  Call CupidsTurn
            End If
            If lblPlayerRole4.Caption = Role Then
                Call CupidsTurn
            End If
            
        'even things out
        End If
    End If
    


'If lblPlayerRole1.Caption = Role Then
    'lblRole.Caption = "Role's Turn: Wolf"
    'lblPlayers.Caption = "Players: " & lblPlayerName1.Caption
    
        'check for more wolves
        'If IsMoreThanOneWolf = True Then
            'If LblPlayerRole2.Caption = Role Then
'End If

'End If
'End If

End Sub

Private Sub cmdPlayer1_Click()


If cmdPlayer1.Visible = True Then
    If fraPlayer1.Visible = False Then
        Call HideFra
        fraPlayer1.Visible = True
    Else
        Call HideFra
    End If
End If

End Sub

Private Sub cmdPlayer2_Click()

If cmdPlayer2.Visible = True Then
    If fraPlayer2.Visible = False Then
        Call HideFra
        fraPlayer2.Visible = True
    Else
        Call HideFra
    End If
End If

End Sub

Private Sub cmdPlayer3_Click()

If cmdPlayer3.Visible = True Then
    If fraPlayer3.Visible = False Then
        Call HideFra
        fraPlayer3.Visible = True
    Else
        Call HideFra
    End If
End If

End Sub

Private Sub cmdPlayer4_Click()

If cmdPlayer4.Visible = True Then
    If fraPlayer4.Visible = False Then
        Call HideFra
        fraPlayer4.Visible = True
    Else
        Call HideFra
    End If
End If

End Sub

Private Sub Command1_Click()
Dim msg, sapi
msg = "It is time for the wolf to wake up."
Set sapi = CreateObject("sapi.spvoice")

sapi.speak msg

' x = inputbox("STUFF HERE!, TEXTBOXNAME HERE")

End Sub

Private Sub Form_Load()

'SAPI###' is where all the voice is done. It makes testing much slower, so only use them later.

With frmMain
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With

Call HideFra

End Sub

Function IsMoreThanOneWolf() As Boolean
Dim wolves As Byte

IsMoreThanOneWolf = False

If lblPlayerRole1.Caption = "Role: Wolf" Then
wolves = wolves + 1
End If

If LblPlayerRole2.Caption = "Role: Wolf" Then
wolves = wolves + 1
End If

If lblPlayerRole3.Caption = "Role: Wolf" Then
wolves = wolves + 1
End If

If lblPlayerRole4.Caption = "Role: Wolf" Then
wolves = wolves + 1
End If

If wolves > 1 Then
    IsMoreThanOneWolf = True
Else
    IsMoreThanOneWolf = False
End If
End Function

Public Sub HideFra()

fraPlayer1.Visible = falses
fraPlayer2.Visible = False
fraPlayer3.Visible = False
fraPlayer4.Visible = False

End Sub
