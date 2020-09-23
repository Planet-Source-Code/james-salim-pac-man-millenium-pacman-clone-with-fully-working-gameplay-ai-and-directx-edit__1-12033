VERSION 5.00
Begin VB.Form LevelProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Level Properties"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4125
   Icon            =   "frmLevelEditProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fLevel 
      Caption         =   "&Level"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtLevelName 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lLevelName 
         Caption         =   "Level Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "LevelProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    txtLevelName = RTrim$(lvledBody.lvlName)
End Sub

Private Sub txtLevelName_Change()
    lvledBody.lvlName = txtLevelName
End Sub

