Attribute VB_Name = "modMainLogic"
Option Explicit

Sub Main()

'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

With frmMain
    .Caption = GAME_LOADING
    .Visible = True
    
    'set the max
    .scrlPlayers.Max = MAX_PLAYERS
    .scrlWolves.Max = MAX_PLAYERS / 2
End With

'Avoid a RTE on setup.
Game.NumberofWolves = 1
Game.ADRole(Role.Villager) = True
Game.ADRole(Role.Wolf) = True

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Public Sub ReportError(ByVal ErrorNumber As Long, ByVal Desc As String)

    MsgBox "A RTE #" & ErrorNumber & " popped up: " & Desc

End Sub

Public Sub AddText(ByVal Text As String)
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

    frmMain.lblInfo.Caption = Text
    
Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Public Sub ActionMSG(ByVal Text As String)
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

If frmActions.Visible = True And ViewActions = True Then
    frmActions.txtInformation.Text = frmActions.txtInformation.Text & vbCrLf & Text
    frmActions.txtInformation.SelLength = Len(frmActions.txtInformation.Text)
End If

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

Public Function RAND()
'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Randomize
RAND = Int((100 - 1 + 1) * Rnd) + 1
If frmActions.Visible = True Then frmActions.lblRAND.Caption = RAND

Exit Function
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Function

Public Sub SetupGame()
Dim Players As Long
Dim FoundRole As Boolean
Dim NumofRoles, ADRoles As Byte
Dim LoopNumber As Long
Dim ADBA(1 To 5) As Boolean
Dim i As Byte
'ADBA = additional role already been assigned

'in case something happens.
If InDebug = True Then On Error GoTo errorhandler:

Players = Game.NumberofPlayers

'Add up the number of roles
NumofRoles = Game.NumberofWolves

'make sure that there are enough people for the roles
ADRoles = 0
If Game.ADRole(Role.Cupid) = True Then ADRoles = ADRoles + 1
If Game.ADRole(Role.Guardian) = True Then ADRoles = ADRoles + 1
If Game.ADRole(Role.Witch) = True Then ADRoles = ADRoles + 1

If Game.NumberofPlayers = 0 Or Game.NumberofWolves = 0 Then
    Call ActionMSG("No players or wolves set.")
    Exit Sub
End If

'if there are too many, exit out.
If ADRoles > Game.NumberofPlayers - Game.NumberofWolves Then
    Call ActionMSG("Too many additional roles.")
    Call AddText("Too many additional roles.")
    Exit Sub
End If

'everything was fine; continue
'NumofRoles = NumofRoles + ADRoles

'assign roles
For i = 1 To Players
    FoundRole = False
    LoopNumber = 0

    Do While FoundRole = False
    
    'try to assign a wolf role
    If Game.NumberofWolves > 0 And FoundRole = False Then
        If RAND / 100 <= (Game.NumberofWolves / Game.NumberofPlayers) Then
            Player(i).Role = Role.Wolf
            Player(i).IsAlive = True
            Player(i).Name = InputBox("This player will become the wolf.", "Player " & i)
            FoundRole = True
            Game.NumberofWolves = Game.NumberofWolves - 1
        End If
    End If
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\
    'try to assign other roles
    '\\\\\\\\\\\\\\\\\\\\\\\\\
    
    If Game.ADRole(Role.Cupid) = True And FoundRole = False Then
        'check to see if its already been assigned
        If ADBA(Role.Cupid) = False Then
            If RAND / 100 <= (1 / Game.NumberofPlayers) Then
                Player(i).Role = Role.Wolf
                Player(i).IsAlive = True
                Player(i).Name = InputBox("This player will become Cupid.", "Player " & i)
                FoundRole = True
                ADBA(Role.Cupid) = True
            End If
        End If
    End If
    
    If Game.ADRole(Role.Guardian) = True And FoundRole = False Then
        'check to see if its already been assigned
        If ADBA(Role.Guardian) = False Then
            If RAND / 100 <= (1 / Game.NumberofPlayers) Then
                Player(i).Role = Role.Guardian
                Player(i).IsAlive = True
                Player(i).Name = InputBox("This player will become the guardian.", "Player " & i)
                ADBA(Role.Guardian) = True
                FoundRole = True
            End If
        End If
    End If
    
    If Game.ADRole(Role.Witch) = True And FoundRole = False Then
        'check to see if its already been assigned
        If ADBA(Role.Witch) = False Then
            If RAND / 100 <= (1 / Game.NumberofPlayers) Then
                Player(i).Role = Role.Witch
                Player(i).IsAlive = True
                Player(i).Name = InputBox("This player will become the witch.", "Player " & i)
                ADBA(Role.Witch) = True
                FoundRole = True
            End If
        End If
    End If
                
    
    'if the loop is going on too long, there's probably no more roles, so set the player to a villager
    If LoopNumber = 30000 And FoundRole = False Then
        Player(i).Role = Role.Villager
        Player(i).IsAlive = True
        Player(i).Name = InputBox("This player will become a villager.", "Player " & i)
        FoundRole = True
    End If
    
    If FoundRole = True Then
        Call ActionMSG("Loop number: " & LoopNumber)
        MsgBox "Call another person to come forward. This msgbox is to prevent other players from seeing the next person or previous person's assigned role."
    End If
    
    LoopNumber = LoopNumber + 1
    DoEvents
    Loop
Next

Call ActionMSG("The game finished loading.")

frmGame.Show
frmMain.Hide
'frmAdminForm.Show
frmGame.cmdNextTurn.Visible = True

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)
End Sub

