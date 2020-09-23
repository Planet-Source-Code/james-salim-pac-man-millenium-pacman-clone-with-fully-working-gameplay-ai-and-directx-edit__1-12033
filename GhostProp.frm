VERSION 5.00
Begin VB.Form GhostProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ghost Properties"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2025
   Icon            =   "GhostProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fGhost 
      Caption         =   "&Ghost"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.TextBox txtGhostSDelay 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtGhostDelay 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtGhostPosY 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtGhostPosX 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   495
      End
      Begin VB.Frame fGhostColour 
         Caption         =   "Colour"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton oGhostRed 
            BackColor       =   &H000000FF&
            Caption         =   "Red"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton oGhostCyan 
            BackColor       =   &H00FFFF00&
            Caption         =   "Cyan"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton oGhostGreen 
            BackColor       =   &H00008000&
            Caption         =   "Green"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton oGhostYellow 
            BackColor       =   &H0000FFFF&
            Caption         =   "Yellow"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Label lGhostSickDelay 
         Caption         =   "Sick Delay"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lGhostDelay 
         Caption         =   "Start Delay"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lGhostPos1 
         Alignment       =   2  'Center
         Caption         =   ","
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lGhostPos 
         Caption         =   "x, y"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
   End
End
Attribute VB_Name = "GhostProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    oGhostRed = True
    RefreshGhostProp
End Sub

Sub RefreshGhostProp()
    Dim TempGhost As LevelGhostChar
    
    If oGhostRed Then lvledpropGhostSelected = 1
    If oGhostCyan Then lvledpropGhostSelected = 2
    If oGhostGreen Then lvledpropGhostSelected = 3
    If oGhostYellow Then lvledpropGhostSelected = 4
    TempGhost = lvledBody.lvlGhost(lvledpropGhostSelected)
    
    txtGhostPosX = TempGhost.X
    txtGhostPosY = TempGhost.Y
    txtGhostDelay = TempGhost.Delay
    txtGhostSDelay = TempGhost.SickDelay
End Sub

Private Sub oGhostCyan_Click()
    RefreshGhostProp
End Sub

Private Sub oGhostGreen_Click()
    RefreshGhostProp
End Sub

Private Sub oGhostRed_Click()
    RefreshGhostProp
End Sub

Private Sub oGhostYellow_Click()
    RefreshGhostProp
End Sub
Private Sub txtGhostDelay_Change()
    lvledBody.lvlGhost(lvledpropGhostSelected).Delay = Val(txtGhostDelay)
End Sub

Private Sub txtGhostPosX_Change()
    lvledBody.lvlGhost(lvledpropGhostSelected).X = Val(txtGhostPosX)
End Sub

Private Sub txtGhostPosY_Change()
    lvledBody.lvlGhost(lvledpropGhostSelected).Y = Val(txtGhostPosY)
End Sub

Private Sub txtGhostSDelay_Change()
    lvledBody.lvlGhost(lvledpropGhostSelected).SickDelay = Val(txtGhostSDelay)
End Sub
