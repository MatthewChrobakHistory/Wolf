VERSION 5.00
Begin VB.Form frmAdminForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAlive 
      Caption         =   "Alive?"
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cmbRole 
      Height          =   315
      ItemData        =   "frmAdminForm.frx":0000
      Left            =   2400
      List            =   "frmAdminForm.frx":0002
      TabIndex        =   4
      Text            =   "Villager"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox lstIndex 
      Height          =   5910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Role:"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadAdminForm()
Dim i As Byte

'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

lstIndex.Clear

For i = 1 To Game.NumberofPlayers
    lstIndex.AddItem i & ": " & Trim$(Player(i).Name)
Next

cmbRole.Clear

For i = 1 To 5
    Select Case i
        Case Role.Witch
            cmbRole.AddItem "Witch", Role.Witch
        Case Role.Cupid
            cmbRole.AddItem "Cupid", Role.Cupid
        Case Role.Guardian
            cmbRole.AddItem "Guardian", Role.Guardian
        Case Role.Wolf
            cmbRole.AddItem "Wolf", Role.Wolf
        Case Role.Villager
            cmbRole.AddItem "Villager", Role.Villager
    End Select
Next

Call lstIndex_Click

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub cmbRole_Change()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Player(i).Role = cmbRole.ListIndex + 1

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub lstIndex_Click()
Dim i As Byte

'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

'set the index, and try to avoid RTE
PlayerIndex = lstIndex.ListIndex + 1
If PlayerIndex = 0 Then PlayerIndex = 1

txtName.Text = Trim$(Player(PlayerIndex).Name)
cmbRole.ListIndex = Player(PlayerIndex).Role - 1
If Player(PlayerIndex).IsAlive = True Then
    chkAlive.Value = 1
Else
    chkAlive.Value = 0
End If

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub txtName_Change()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Player(PlayerIndex).Name = Trim$(txtName.Text)

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub
