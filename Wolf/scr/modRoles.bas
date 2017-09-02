Attribute VB_Name = "modRoles"
Option Explicit

'character values
'//CUPID
Public CupidSelect(1 To 2) As Byte

'//WOLF
Public WolfVictim As Byte

'//WITCH
Public WitchTurn As Byte
Public HealPotion As Boolean
Public HealIndex As Byte
Public KillPotion As Boolean
Public KillIndex As Byte


'//GUARDIAN
Public GuardianSave As Byte


Public Sub PlayTurn(ByVal RoleNumber As Byte)
'in case something happens
If InDebug = True Then On Error GoTo errorhandler:

Select Case RoleNumber
    Case Role.Cupid
        With frmGame
            .fraRoles.Visible = True
            .fraRoles.Caption = "Select Data: Cupid (1)"
            Call PopulateList
        End With
    Case Role.Guardian
        With frmGame
            .fraRoles.Visible = True
            .fraRoles.Caption = "Select Data: Guardian"
            Call PopulateList
        End With
    Case Role.Witch
        With frmGame
            .fraRoles.Visible = True
            .fraRoles.Caption = "Select Data: Witch (SAVE)"
            If HealPotion = True Then
                .cmdSelect.Visible = False
            Else
                .cmdSelect.Visible = True
            End If
            'If WolfVictim = 0 Then WolfVictim = 1
            .lblAdiInfo.Caption = "Tell the Witch that " & Trim$(Player(WolfVictim).Name) & " is dead."
            WitchTurn = 1
            .cmdPass.Visible = True
            Call PopulateList
        End With
    Case Role.Wolf
        With frmGame
            .fraRoles.Visible = True
            .fraRoles.Caption = "Select Data: Wolf"
            Call PopulateList
        End With
    Case Role.Villager
        Call NightTimeLogic
        frmGame.fraDayTime.Visible = True
        Call PopulateList
End Select

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Public Sub PopulateList()
Dim i As Long
Dim State As String

'in case something happens
If InDebug = True Then On Error GoTo errorhandler:

With frmGame
    If .lstIndex.Visible = True Then
        .lstIndex.Clear
        For i = 1 To Game.NumberofPlayers
            If Player(i).IsAlive = True Then
                .lstIndex.AddItem (i & ": " & Trim$(Player(i).Name))
            Else
                .lstIndex.AddItem (i & ": DEAD")
            End If
        Next
    End If
    
    If .lstDayIndex.Visible = True Then
        .lstDayIndex.Clear
        For i = 1 To Game.NumberofPlayers
            If Player(i).IsAccused = True Then
                State = "ACCUSED"
            Else
                State = vbNullString
            End If
            If Player(i).IsAlive = True Then
                .lstDayIndex.AddItem (i & ": " & Trim$(Player(i).Name & " " & State))
            Else
                .lstDayIndex.AddItem (i & ": DEAD")
            End If
        Next
    End If
End With
            

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub
