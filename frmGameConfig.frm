VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Configuration"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin MSComctlLib.Slider sPoint 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   30
      SelStart        =   1
      TickFrequency   =   5
      Value           =   1
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin MSComctlLib.Slider sSpeed 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Min             =   -20
      Max             =   20
      TickFrequency   =   5
   End
   Begin MSComctlLib.Slider sLevel 
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin MSComctlLib.Slider sLife 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
   End
   Begin MSComctlLib.Slider sAggressive 
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Min             =   -50
      Max             =   50
      TickFrequency   =   10
   End
   Begin VB.Label txtAggressive 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Ghost Aggressiveness"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label txtSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label txtLevel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label txtPoint 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label txtLife 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Fast"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Slow"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblSpeed 
      Alignment       =   1  'Right Justify
      Caption         =   "&Speed Adjustment"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting L&evel"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblPoint 
      Alignment       =   1  'Right Justify
      Caption         =   "&Points for bonus life"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblLife 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of &Lives"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmGameConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    AdjustingSpeed = sSpeed.Value * -1
    PacLevelStart = sLevel.Value
    PacScoretoLife = sPoint.Value * 100
    PacLifeStart = sLife.Value
    GhostAgressivity = sAggressive.Value
    Unload Me
End Sub

Private Sub Form_Load()
    sSpeed.Value = AdjustingSpeed * -1
    sLevel.Max = MaxLevelNo
    sLevel.Value = PacLevelStart
    sPoint.Value = PacScoretoLife / 100
    sLife.Value = PacLifeStart
    sAggressive.Value = GhostAgressivity
    
    sLevel_Change
    sLife_Change
    sPoint_Change
    sSpeed_Change
    sAggressive_Change
End Sub

Private Sub sAggressive_Change()
    txtAggressive = sAggressive.Value
End Sub

Private Sub sLevel_Change()
    txtLevel.Caption = LTrim$(Str$(sLevel.Value))
End Sub

Private Sub sLife_Change()
    txtLife.Caption = LTrim$(Str$(sLife.Value))
End Sub

Private Sub sPoint_Change()
    txtPoint.Caption = Format(sPoint.Value * 1000, "###,###,###")
End Sub

Private Sub sSpeed_Change()
    txtSpeed.Caption = LTrim$(Str$(sSpeed.Value))
End Sub
