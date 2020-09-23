VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLevelEditSchemes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Schemes..."
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmLevelEditSchemes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox schBack 
      Height          =   6300
      Left            =   0
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   416
      TabIndex        =   0
      Top             =   0
      Width           =   6300
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wall 2"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wall 1"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Food"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   1920
         Width           =   855
      End
      Begin VB.Image schWall2 
         Height          =   330
         Left            =   4320
         Top             =   1560
         Width           =   330
      End
      Begin VB.Image schWall1 
         Height          =   330
         Left            =   2880
         Top             =   1560
         Width           =   330
      End
      Begin VB.Image schFood 
         Height          =   330
         Left            =   1440
         Top             =   1560
         Width           =   330
      End
   End
   Begin MSComDlg.CommonDialog DlgFile 
      Left            =   0
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Choose Schemes..."
   End
   Begin VB.Label Label4 
      Caption         =   "NOTE: Adjust schemes by clicking on the picture"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6360
      Width           =   5295
   End
End
Attribute VB_Name = "frmLevelEditSchemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    schBack.Picture = LoadPicture(fpImage + "schemes\" + LTrim$(Str$(lvledBody.lvlScheme.Back)) + "_back.bmp")
    schFood.Picture = frmLevelEditor.SchemeList.ListImages(1).ExtractIcon
    schWall1.Picture = frmLevelEditor.SchemeList.ListImages(3).ExtractIcon
    schWall2.Picture = frmLevelEditor.SchemeList.ListImages(4).ExtractIcon

    DlgFile.InitDir = fpImage + "schemes\"
End Sub

Private Sub schBack_Click()
    On Error GoTo ErrorProc
    DlgFile.Filter = "*_back.bmp (Background Scheme)|?_back.bmp"
    DlgFile.ShowOpen
    If LCase(Right$(DlgFile.FileName, 9)) = "_back.bmp" Then
        lvledBody.lvlScheme.Back = Val(Left$(Right$(DlgFile.FileName, 10), 1))
        lvledRefreshSchemes
        schBack.Picture = LoadPicture(DlgFile.FileName)
    End If
ErrorProc:
End Sub

Private Sub schFood_Click()
On Error GoTo ErrorProc
    DlgFile.Filter = "*_food.bmp (Food Scheme)|?_food.bmp"
    DlgFile.ShowOpen
    If LCase(Right$(DlgFile.FileName, 9)) = "_food.bmp" Then
        lvledBody.lvlScheme.Food = Val(Left$(Right$(DlgFile.FileName, 10), 1))
        lvledRefreshSchemes
        schFood.Picture = frmLevelEditor.SchemeList.ListImages(1).ExtractIcon
    End If
ErrorProc:
End Sub

Private Sub schWall1_Click()
On Error GoTo ErrorProc
    DlgFile.Filter = "*_wall.bmp (Wall #1 Scheme)|?_wall.bmp"
    DlgFile.ShowOpen
    If LCase(Right$(DlgFile.FileName, 9)) = "_wall.bmp" Then
        lvledBody.lvlScheme.Wall1 = Val(Left$(Right$(DlgFile.FileName, 10), 1))
        lvledRefreshSchemes
        schWall1.Picture = frmLevelEditor.SchemeList.ListImages(3).ExtractIcon
    End If
ErrorProc:
End Sub

Private Sub schWall2_Click()
On Error GoTo ErrorProc
    DlgFile.Filter = "*_wall2.bmp (Wall #2 Scheme)|?_wall2.bmp"
    DlgFile.ShowOpen
    If LCase(Right$(DlgFile.FileName, 10)) = "_wall2.bmp" Then
        lvledBody.lvlScheme.Wall2 = Val(Left$(Right$(DlgFile.FileName, 11), 1))
        lvledRefreshSchemes
        schWall2.Picture = frmLevelEditor.SchemeList.ListImages(4).ExtractIcon
    End If
ErrorProc:
End Sub
