VERSION 5.00
Begin VB.Form frmAdminLogin 
   Caption         =   "Admin Login"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLogin 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "login"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Password:"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Username:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()

If txtUsername.Text <> "" Then
    MsgBox "Incorrect: Redirecting back to game."
Else
    frmMain.fraAdminPanel.Visible = True
End If

End Sub

Private Sub Form_Load()

With frmAdminLogin
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With

End Sub
