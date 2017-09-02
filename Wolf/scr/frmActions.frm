VERSION 5.00
Begin VB.Form frmActions 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wolf: Game Actions"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInformation 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000004&
      Height          =   2295
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblRAND 
      BackColor       =   &H80000007&
      Caption         =   "Random Number: 000"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
End
Attribute VB_Name = "frmActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblRAND_Click()

Call RAND

End Sub
