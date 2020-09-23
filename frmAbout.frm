VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About PacMan Millenium"
   ClientHeight    =   5220
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3765
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3602.937
   ScaleMode       =   0  'User
   ScaleWidth      =   3535.53
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox iconPacDog 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2700
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      ToolTipText     =   "PacDog"
      Top             =   4410
      Width           =   480
   End
   Begin VB.PictureBox iconPacMan 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   720
      Picture         =   "frmAbout.frx":1194
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      ToolTipText     =   "PacMan"
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox iconPacCat 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   270
      Picture         =   "frmAbout.frx":1A5E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      ToolTipText     =   "PacCat"
      Top             =   2880
      Width           =   480
   End
   Begin VB.PictureBox iconJunior 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":2328
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      ToolTipText     =   "Junior"
      Top             =   2040
      Width           =   480
   End
   Begin VB.PictureBox iconMsPacMan 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":2BF2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      ToolTipText     =   "Ms PacMan"
      Top             =   1080
      Width           =   480
   End
   Begin VB.PictureBox iconSue 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   630
      Picture         =   "frmAbout.frx":34BC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      ToolTipText     =   "Sue"
      Top             =   4365
      Width           =   480
   End
   Begin VB.PictureBox iconBlinky 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2610
      Picture         =   "frmAbout.frx":3D86
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      ToolTipText     =   "Blinky"
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox iconRed 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3105
      Picture         =   "frmAbout.frx":4650
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      ToolTipText     =   "Red"
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox iconPink 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3120
      Picture         =   "frmAbout.frx":4F1A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      ToolTipText     =   "Pink"
      Top             =   2040
      Width           =   480
   End
   Begin VB.PictureBox iconInky 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3120
      Picture         =   "frmAbout.frx":57E4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      ToolTipText     =   "Inky"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) James Salim, 2000-2002"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   14
      Top             =   4005
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click anywhere inside the form to close the About box."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   855
      TabIndex        =   13
      Top             =   3465
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":60AE
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   840
      TabIndex        =   12
      Top             =   1605
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program and Designed by James Salim."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   1125
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PacMan Millenium"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   765
      Width           =   2175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub
