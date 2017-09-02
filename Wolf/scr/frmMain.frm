VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wolf: Setup"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "General Information"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CheckBox chkActions 
         Caption         =   "View Actions?"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Frame fraRoles 
         Caption         =   "Additional Roles"
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   2535
         Begin VB.CheckBox chkCupid 
            Caption         =   "Cupid"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox chkGuardian 
            Caption         =   "Guardian"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chkWitch 
            Caption         =   "Witch"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Debug?"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   1095
      End
      Begin VB.HScrollBar scrlWolves 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   1
         TabIndex        =   4
         Top             =   1200
         Value           =   1
         Width           =   2175
      End
      Begin VB.HScrollBar scrlPlayers 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   1
         TabIndex        =   1
         Top             =   600
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label lblWolves 
         Caption         =   "Number of Wolves: 1"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblPlayers 
         Caption         =   "Players: 1"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "Homescreen Loaded"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkActions_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

ViewActions = chkActions.Value

If ViewActions = True And frmActions.Visible = False Then frmActions.Show

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub chkCupid_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Game.ADRole(Role.cupid) = chkCupid.Value

Select Case Game.ADRole(Role.cupid)
    Case True
        Call ActionMSG("Cupid was enabled.")
    Case Else
        Call ActionMSG("Cupid was disabled.")
End Select

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub chkDebug_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

InDebug = chkDebug.Value

Call ActionMSG("Debug is set to " & chkDebug.Value)

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub chkGuardian_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Game.ADRole(Role.Guardian) = chkGuardian.Value

Select Case Game.ADRole(Role.Guardian)
    Case True
        Call ActionMSG("The Guardian was enabled.")
    Case Else
        Call ActionMSG("The Guardian was disabled.")
End Select

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub chkWitch_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Game.ADRole(Role.Witch) = chkWitch.Value

Select Case Game.ADRole(Role.Witch)
    Case True
        Call ActionMSG("The Witch was enabled.")
    Case Else
        Call ActionMSG("The Witch was disabled.")
End Select

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub cmdPlay_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Call ActionMSG("The game is now loading...")
Call SetupGame

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub


Private Sub Form_Unload(Cancel As Integer)
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

End

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub scrlPlayers_Change()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

If scrlPlayers.Value / 2 < scrlWolves Then
    Call AddText("There are too many wolves! Increase the amount of players, or decrease the amount of wolves!")
    scrlPlayers.Value = Game.NumberofPlayers
    Exit Sub
End If

lblPlayers.Caption = "Players: " & scrlPlayers.Value
Game.NumberofPlayers = scrlPlayers.Value

Call ActionMSG("Number of players was set to " & Game.NumberofPlayers)

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub scrlWolves_Change()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

If scrlPlayers.Value / 2 < scrlWolves Then
    Call AddText("There are too many wolves! Increase the amount of players, or decrease the amount of wolves!")
    scrlWolves.Value = Game.NumberofWolves
    Exit Sub
End If

lblWolves.Caption = "Number of Wolves: " & scrlWolves.Value
Game.NumberofWolves = scrlWolves.Value

Call ActionMSG("Number of wolves was set to " & Game.NumberofWolves)

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub
