VERSION 5.00
Begin VB.Form PacProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PacMan Properties"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   Icon            =   "frmLevelEditPacProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fPac 
      Caption         =   "&Pac-Man"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtPacPosX 
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtPacPosY 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Frame sfDirection 
         Caption         =   "Direction"
         Height          =   1335
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton oPacUp 
            Caption         =   "&Up"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton oPacDn 
            Caption         =   "&Down"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton oPacLf 
            Caption         =   "&Left"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton oPacRg 
            Caption         =   "&Right"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.TextBox txtPacDelay 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtPacDDelay 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame sfTime 
         Caption         =   "Time"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   2895
         Begin VB.TextBox txtShieldTime 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtDrunkTime 
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lShield 
            Caption         =   "Shield"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lDrunk 
            Caption         =   "Drunk"
            Height          =   255
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lPacPos 
         Caption         =   "x, y"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lDelay 
         Caption         =   "Start Delay"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lPacPos1 
         Alignment       =   2  'Center
         Caption         =   ","
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lSickDelay 
         Caption         =   "Drunk Delay"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
   End
End
Attribute VB_Name = "PacProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '--Pac--
    With lvledBody.lvlPac
        txtPacPosX = .X
        txtPacPosY = .Y
        txtPacDelay = .StartDelay
        txtPacDDelay = .DrunkDelay
        txtShieldTime = .ProtectTime
        txtDrunkTime = .DrunkTime
        If .xDir = -1 Then oPacLf.Value = True
        If .xDir = 1 Then oPacRg.Value = True
        If .yDir = -1 Then oPacUp.Value = True
        If .yDir = 1 Then oPacDn.Value = True
    End With
End Sub

Private Sub oPacDn_Click()
    lvledBody.lvlPac.xDir = 0
    lvledBody.lvlPac.yDir = 1
End Sub

Private Sub oPacLf_Click()
    lvledBody.lvlPac.xDir = -1
    lvledBody.lvlPac.yDir = 0
End Sub

Private Sub oPacRg_Click()
    lvledBody.lvlPac.xDir = 1
    lvledBody.lvlPac.yDir = 0
End Sub

Private Sub oPacUp_Click()
    lvledBody.lvlPac.xDir = 0
    lvledBody.lvlPac.yDir = -1
End Sub

Private Sub txtDrunkTime_Change()
    lvledBody.lvlPac.DrunkTime = Val(txtDrunkTime)
End Sub

Private Sub txtPacDDelay_Change()
    lvledBody.lvlPac.DrunkDelay = Val(txtPacDDelay)
End Sub

Private Sub txtPacDelay_Change()
    lvledBody.lvlPac.StartDelay = Val(txtPacDelay)
End Sub

Private Sub txtPacPosX_Change()
    lvledBody.lvlPac.X = Val(txtPacPosX)
End Sub

Private Sub txtPacPosY_Change()
    lvledBody.lvlPac.Y = Val(txtPacPosY)
End Sub

Private Sub txtShieldTime_Change()
    lvledBody.lvlPac.ProtectTime = Val(txtShieldTime)
End Sub

