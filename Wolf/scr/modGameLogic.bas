Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub NextTurn()
' just in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Call CheckForWin

If AreWolvesAlive = True Then
    Select Case FindNextRole
        Case Role.Cupid
            frmGame.lblCurRoleTurn = "Current Role Turn: " & GetRoleName(Role.Cupid)
            frmGame.lblCurTime.Caption = "Current Time: Night"
            Call PlayTurn(Role.Cupid)
        Case Role.Wolf
            frmGame.lblCurRoleTurn.Caption = "Current Role Turn: " & GetRoleName(Role.Wolf)
            frmGame.lblCurTime.Caption = "Current Time: Night"
            Call PlayTurn(Role.Wolf)
        Case Role.Witch
            frmGame.lblCurRoleTurn.Caption = "Current Role Turn: " & GetRoleName(Role.Witch)
            Call PlayTurn(Role.Witch)
        Case Role.Guardian
            frmGame.lblCurRoleTurn.Caption = "Current Role Turn: " & GetRoleName(Role.Guardian)
            Call PlayTurn(Role.Guardian)
        Case Role.Villager
            frmGame.lblCurRoleTurn.Caption = "Current Role Turn: " & GetRoleName(Role.Villager)
            frmGame.lblCurTime.Caption = "Current Time: Day"
            frmGame.cmdNextTurn.Visible = False
            Call PlayTurn(Role.Villager)
            'turn to day
    End Select
    
    CurRoleOn = FindNextRole
    
Else
    MsgBox "VILLAGERS WIN"
    Call EndGame
End If

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Public Sub EndGame()
' just in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

End

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Public Function AreWolvesAlive() As Boolean
Dim i As Long

' just in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

For i = 1 To MAX_PLAYERS
    If Player(i).IsAlive = True Then
        If Player(i).Role = Role.Wolf Then
            AreWolvesAlive = True
        End If
    End If
Next

Exit Function
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Function

Public Function FindNextRole() As Byte
' just in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

'cupid
If CurRoleOn = 0 And Game.ADRole(Role.Cupid) = True Then
    FindNextRole = Role.Cupid
    Exit Function
End If

'wolf
If CurRoleOn = Role.Cupid Or CurRoleOn = Role.Villager Or CurRoleOn = 0 And Game.ADRole(Role.Cupid) = False Then
    If Game.ADRole(Role.Wolf) = True Then
        FindNextRole = Role.Wolf
        Exit Function
    End If
End If

'witch
If CurRoleOn = Role.Wolf And Game.ADRole(Role.Witch) = True Then
    FindNextRole = Role.Witch
    Exit Function
End If

'guardian
If Game.ADRole(Role.Guardian) = True Then
    If CurRoleOn = Role.Witch Or CurRoleOn = Role.Wolf And Game.ADRole(Role.Witch) = False Then
        FindNextRole = Role.Guardian
        Exit Function
    End If
End If

'villager/turn to day
If CurRoleOn = Role.Guardian Or CurRoleOn = Role.Witch Or CurRoleOn = Role.Wolf Then
    FindNextRole = Role.Villager
    Exit Function
End If
    
Exit Function
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Function

Public Function GetRoleName(ByVal RoleNumber As Byte) As String
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Select Case RoleNumber
    Case Role.Cupid
        GetRoleName = "Cupid"
    Case Role.Guardian
        GetRoleName = "Guardian"
    Case Role.Villager
        GetRoleName = "Villager"
    Case Role.Witch
        GetRoleName = "Witch"
    Case Role.Wolf
        GetRoleName = "Wolf"
End Select

Exit Function
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Function

Public Sub CheckForWin()
Dim i As Long
Dim Villagers As Long
Dim Wolves As Long

For i = 1 To MAX_PLAYERS
    If Player(i).Role = Role.Wolf Then
        Wolves = Wolves + 1
    Else
        Villagers = Villagers + 1
    End If
Next

End Sub
