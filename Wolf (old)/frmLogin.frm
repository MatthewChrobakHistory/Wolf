VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLogin 
      Caption         =   "Info"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   600
         ScaleHeight     =   202
         ScaleMode       =   0  'User
         ScaleWidth      =   248.966
         TabIndex        =   4
         Top             =   1080
         Width           =   4725
      End
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H8000000B&
         Caption         =   "Play"
         Height          =   915
         Left            =   1440
         TabIndex        =   3
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox txtPlayers 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "How many players?"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFatBoy_Click()

If chkInfoMode.Value = 1 Then
    MsgBox ("The Fat Boy with a Backpack is a classic role in Wolf. When the Fat Boy dies, the person on the right of him or her is eaten by the Fat Boy.")
End If

End Sub

Private Sub cmdExit_Click()

If fraMinorRoles.Visible = False Then
    fraRoles.Visible = False
Else
    fraMinorRoles.Visible = False
End If

End Sub

Private Sub cmdMinorRoles_Click()

fraMinorRoles.Visible = True

End Sub

Private Sub cmdPlay_Click()

Call SetPlayerNames

End Sub

Private Sub cmdRoles_Click()

fraRoles.Visible = True

End Sub

Private Sub Form_Load()

With frmLogin
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With

cmdPlay.Visible = False
Picture1.Picture = LoadPicture(App.Path & "\Artwork\wolf.jpg")

End Sub

Private Sub txtPlayers_Change()

If IsNumeric(txtPlayers.Text) = True Then
    If txtPlayers.Text > 0 Then
        cmdPlay.Visible = True
    End If
Else
    cmdPlay.Visible = False
End If



End Sub

Private Sub txtWolves_Change()
If IsNumeric(txtPlayers.Text) = True Then
    If txtPlayers.Text > 0 Then
        If IsNumeric(txtWolves.Text) = True Then
            If txtWolves.Text > 0 Then
                cmdPlay.Visible = True
            End If
        End If
    End If
Else
    cmdPlay.Visible = False
End If
End Sub

Public Function SetPlayerNames()
Dim Name As String
Dim Role As String
Dim Players As Byte

frmMain.Show
frmMain.Hide

Players = txtPlayers.Text

If Players <> 0 Then
'name
Name = InputBox("Enter the name of the first player.", "Player Names")
frmMain.lblPlayerName1.Caption = "Name: " & Name
frmMain.cmdPlayer1.Caption = Name
'role
Role = InputBox("Enter " & Name & "'s role.", Name & "'s role")
frmMain.lblPlayerRole1.Caption = "Role: " & Role
'other stuff
frmMain.lblPlayerAlive1.Caption = "Alive: Yes"
frmMain.cmdPlayer1.Visible = True
Players = Players - 1
End If

If Players <> 0 Then
'name
Name = InputBox("Enter the name of the second player.", "Player Names")
frmMain.lblPlayerName2.Caption = "Name: " & Name
frmMain.cmdPlayer2.Caption = Name
'role
Role = InputBox("Enter " & Name & "'s role.", Name & "'s role")
frmMain.LblPlayerRole2.Caption = "Role: " & Role
frmMain.lblPlayerAlive2.Caption = "Alive: Yes"
frmMain.cmdPlayer2.Visible = True
Players = Players - 1
End If

If Players <> 0 Then
Name = InputBox("Enter the name of the third player.", "Player Names")
frmMain.lblPlayerName3.Caption = "Name: " & Name
frmMain.LblPlayerAlive3.Caption = "Alive: Yes"
frmMain.cmdPlayer3.Caption = Name
'role
Role = InputBox("Enter " & Name & "'s role.", Name & "'s role")
frmMain.lblPlayerRole3.Caption = "Role: " & Role
frmMain.cmdPlayer3.Visible = True
Players = Players - 1
End If

If Players <> 0 Then
Name = InputBox("Enter the name of the fourth player.", "Player Names")
frmMain.lblPlayerName4.Caption = "Name: " & Name
frmMain.lblPlayerAlive4.Caption = "Alive: Yes"
frmMain.cmdPlayer4.Caption = Name
'role
Role = InputBox("Enter " & Name & "'s role.", Name & "'s role")
frmMain.lblPlayerRole1.Caption = "Role: " & Role
frmMain.cmdPlayer4.Visible = True
Players = Players - 1
End If

frmMain.Show
frmLogin.Hide

End Function

Private Function IsRoleAvailable()
Dim fatboy As Boolean
Dim witch As Boolean
Dim guardian As Boolean

If frmLogin.chkFatBoy.Value = 1 Then fatboy = True


End Function
