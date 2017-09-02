VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wolf: Game"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDayTime 
      Caption         =   "Daytime Panel"
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   195
         Left            =   4080
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdDayEnd 
         Caption         =   "End Day"
         Height          =   555
         Left            =   2880
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
      End
      Begin VB.HScrollBar scrlVotes 
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkIsAccused 
         Caption         =   "is accused?"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox lstDayIndex 
         Height          =   2595
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblVotes 
         Caption         =   "Votes: None"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblWhoIsDead 
         Caption         =   "Who is dead?: Nobody"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame fraRoles 
      Caption         =   "Select Data: None"
      Height          =   3855
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton cmdPass 
         Caption         =   "Pass"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ListBox lstIndex 
         Height          =   2400
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblAdiInfo 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdNextTurn 
      Caption         =   "START GAME"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame fraGame 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.Label lblCurTime 
         Caption         =   "Current Time: Day"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCurRoleTurn 
         Caption         =   "Current Role Turn: None"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDayEnd_Click()
fraDayTime.Visible = False
Call NextTurn
End Sub

Private Sub cmdNextTurn_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

cmdNextTurn.Visible = False
Call NextTurn

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub cmdPass_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

If CurRoleOn = Role.Witch Then
    lblAdiInfo.Caption = vbNullString
    If WitchTurn = 1 Then
        WitchTurn = 2
        fraRoles.Caption = "Select Data: Witch (KILL)"
        If KillPotion = True Then
            cmdSelect.Visible = False
        Else
            cmdSelect.Visible = True
        End If
    Else
        WitchTurn = 0
        fraRoles.Visible = False
        cmdPass.Visible = False
        cmdSelect.Visible = True
        Call NextTurn
    End If
End If

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub cmdSelect_Click()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

If Player(lstIndex.ListIndex + 1).IsAlive = False Then Exit Sub

'selecting a lover
If CurRoleOn = Role.Cupid Then
    If CupidSelect(1) = 0 Then
        CupidSelect(1) = lstIndex.ListIndex + 1
        fraRoles.Caption = "Select Data: Cupid (2)"
    ElseIf CupidSelect(2) = 0 And CupidSelect(1) <> lstIndex.ListIndex + 1 Then
        CupidSelect(2) = lstIndex.ListIndex + 1
        fraRoles.Visible = False
        Call NextTurn
        Exit Sub
    End If
End If

'selecting someone to save
If CurRoleOn = Role.Guardian Then
    GuardianSave = lstIndex.ListIndex + 1
    Call NextTurn
    fraRoles.Visible = False
    Exit Sub
End If

'witch
If CurRoleOn = Role.Witch Then
    If WitchTurn = 1 Then
        HealIndex = lstIndex.ListIndex + 1
        HealPotion = True
        WitchTurn = 2
        fraRoles.Caption = "Select Data: Witch (KILL)"
        If KillPotion = True Then
            cmdSelect.Visible = False
        Else
            cmdSelect.Visible = True
        End If
        lblAdiInfo.Caption = vbNullString
    ElseIf WitchTurn = 2 Then
        lblAdiInfo.Caption = vbNullString
        KillIndex = lstIndex.ListIndex + 1
        KillPotion = True
        cmdPass.Visible = False
        cmdSelect.Visible = True
        fraRoles.Visible = False
        Call NextTurn
        Exit Sub
    End If
End If

'wolf
If CurRoleOn = Role.Wolf Then
    WolfVictim = lstIndex.ListIndex + 1
    fraRoles.Visible = False
    Call NextTurn
    Exit Sub
End If

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Private Sub Command2_Click()
PopulateList
End Sub

Private Sub Form_Load()
If InDebug = True Then On Error GoTo errorhandler:

scrlVotes.Max = MAX_PLAYERS

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

