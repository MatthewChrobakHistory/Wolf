Attribute VB_Name = "modDayTime"
Option Explicit

Public Sub NightTimeLogic()
Dim i As Byte
Dim CupidsDied As Boolean
Dim WolfDied As Boolean
Dim WitchDied As Boolean
Dim PlayerDied(1 To MAX_PLAYERS) As Boolean
Dim TheDead As String
'CupidSelect(1 to 2) as byte
'WolfVictim as byte
'WitchTurn as byte
'HealPotion as boolean
'HealIndex as byte
'KillPotion as boolean
'KillIndex as byte
'GuardianSave as byte

If InDebug = True Then On Error GoTo errorhandler:

'kill the wolf's victim
Player(WolfVictim).IsAlive = False
PlayerDied(WolfVictim) = True

'see if we can save the player
If HealIndex = WolfVictim Or GuardianSave = WolfVictim Then
    Player(WolfVictim).IsAlive = True
End If

'witch's kill
If KillIndex > 0 Then
    Player(KillIndex).IsAlive = False
    PlayerDied(KillIndex) = True
End If

' cupids
If Player(CupidSelect(1)).IsAlive = False Or Player(CupidSelect(2)).IsAlive = False Then
    Player(CupidSelect(1)).IsAlive = False
    Player(CupidSelect(2)).IsAlive = False
    PlayerDied(CupidSelect(1)) = True
    PlayerDied(CupidSelect(2)) = True
End If

'handle the messages
TheDead = vbNullString

'set the message
For i = 1 To MAX_PLAYERS
    If PlayerDied(i) = True Then
        TheDead = TheDead & ", " & Trim$(Player(i).Name)
        PlayerDied(i) = False
    End If
Next

If TheDead = vbNullString Then TheDead = "Nobody"

frmGame.lblWhoIsDead.Caption = "Who died?: " & TheDead

WolfVictim = 0
WitchTurn = 0
HealIndex = 0
KillIndex = 0
GuardianSave = 0

Exit Sub
errorhandler:
    Call ReportError(Err.Number, Err.Description)

End Sub
