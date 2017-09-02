Attribute VB_Name = "modTypes"
Option Explicit

Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Game As TempGameRec

Private Type PlayerRec
    Name As String
    IsAlive As Boolean
    Role As Byte
    
    'voting
    IsAccused As Boolean
    Votes As Byte
End Type

'this rec is used only for setting roles
Private Type TempGameRec
    NumberofPlayers As Byte
    NumberofWolves As Byte
    
    'additional roles
    ADRole(1 To 5) As Boolean
End Type

'\\\\\\\\\\\\
'ENUMERATIONS
'\\\\\\\\\\\\

Public Enum Role
    Witch = 1
    Guardian
    Cupid
    Wolf
    Villager ' 5
End Enum
