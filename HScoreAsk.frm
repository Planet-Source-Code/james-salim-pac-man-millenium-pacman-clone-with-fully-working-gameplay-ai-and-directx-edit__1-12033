VERSION 5.00
Begin VB.Form HScoreAsk 
   Caption         =   "Congratulations!"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   Icon            =   "HScoreAsk.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lGameOver 
      Alignment       =   2  'Center
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lCongratulations 
      Caption         =   $"HScoreAsk.frx":08CA
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "HScoreAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ReturnName = txtName.Text
    ReplySent = True
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Unload Me
    End If
End Sub
