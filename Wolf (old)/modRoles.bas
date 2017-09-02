Attribute VB_Name = "modRoles"
Public Sub CupidsTurn()
Dim FirstLover As String
Dim SecondLover As String
Dim FirstLoverName As String
Dim SecondLoverName As String
Dim msg, sapi
Dim Useless As String
Set sapi = CreateObject("sapi.spvoice")

With frmMain
.lblRole.Caption = "Role's Turn: Cupid"
.lblPlayers.Caption = "Player: " & .lblPlayerName1.Caption
                
'name of first lover
    FirstLover = InputBox("Enter the name of the first player to fall in love.", "Cupid's turn")
    FirstLoverName = FirstLover
    FirstLover = "Name: " & FirstLover
                
    'find the first lover and make him/her fall in love
                                
    If .lblPlayerName1.Caption = FirstLover Then
        .lblInLove1.Caption = "In Love: Yes"
    End If
                
    If .lblPlayerName2.Caption = FirstLover Then
        .lblInLove2.Caption = "In Love: Yes"
    End If
                
    If .lblPlayerName3.Caption = FirstLover Then
        .lblInLove3.Caption = "In Love: Yes"
    End If
                
    If .lblPlayerName4.Caption = FirstLover Then
        .lblInLove4.Caption = "In Love: Yes"
    End If
                
    'name of the second lover
                
    SecondLover = InputBox("Enter the name of the second player to fall in love.", "Cupid's turn")
    SecondLoverName = SecondLover
    SecondLover = "Name: " & SecondLover
                
    'find the second lover and make him/her fall in love
                                
    If .lblPlayerName1.Caption = SecondLover Then
        .lblInLove1.Caption = "In Love: Yes"
    End If
                
    If .lblPlayerName2.Caption = SecondLover Then
        .lblInLove2.Caption = "In Love: Yes"
    End If
                
    If .lblPlayerName3.Caption = SecondLover Then
        .lblInLove3.Caption = "In Love: Yes"
    End If
                
    If .lblPlayerName4.Caption = SecondLover Then
        .lblInLove4.Caption = "In Love: Yes"
    End If
    
    msg = "I ask cupid to please close his or her eyes."
sapi.speak msg
    
    Useless = InputBox("Pleasy verify that cupid closed his or her eyes. Write anything to continue...")
    
    msg = "The monitor will now go around the room. He or her will bopp the two people who just fell in love. I ask that when the monitor bopps you, you open your eyes."
sapi.speak msg
    
    
    End With
    
    

End Sub

