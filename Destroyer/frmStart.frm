VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destroyer"
   ClientHeight    =   1275
   ClientLeft      =   1755
   ClientTop       =   2025
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   175
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   585
      Left            =   1350
      TabIndex        =   1
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label lblTopScore 
      Caption         =   "Top Score:"
      Height          =   495
      Left            =   30
      TabIndex        =   3
      Top             =   1050
      Width           =   2565
   End
   Begin VB.Label lblKeys 
      Caption         =   "Press right to shoot right.                  Press down to shoot down."
      Height          =   435
      Left            =   30
      TabIndex        =   2
      Top             =   630
      Width           =   2505
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type player
 name As String * 15
 score As Integer
End Type

Public recordlength As Long
Private BlankPlayer As player
Public filename As String

Private Sub cmdExit_Click()
 Unload Me
 Unload frmGame
End Sub

Private Sub cmdStart_Click()
 Unload Me
 frmGame.Show
End Sub

Private Sub Form_Load()
 recordlength = LenB(BlankPlayer)
 filename = App.Path + "/score.log"
 
 Open filename For Random Access Read Write As #1 Len = recordlength
 Get #1, 1, BlankPlayer
 Close #1
 
 lblTopScore.Caption = "Top Score: " + Str(BlankPlayer.score) + " by " + BlankPlayer.name
End Sub
