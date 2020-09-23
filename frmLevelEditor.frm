VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLevelEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pac Man Millenium Level Editor"
   ClientHeight    =   7260
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8205
   Icon            =   "frmLevelEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   1005
      ButtonWidth     =   979
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Food"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Shield"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wall 1"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wall 2"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7005
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "18, 18"
            TextSave        =   "18, 18"
            Object.ToolTipText     =   "Mouse Position on Level Surfaces"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView lvlView 
      Height          =   6270
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   11060
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList SchemeList 
      Left            =   0
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin VB.PictureBox lvlSurface 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   1920
      ScaleHeight     =   418
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   418
      TabIndex        =   0
      Top             =   720
      Width           =   6270
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "&Level"
      Begin VB.Menu lvlSelect 
         Caption         =   "&Select"
      End
      Begin VB.Menu lvlSave 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu lvlInsert 
         Caption         =   "&Insert"
         Begin VB.Menu insAbove 
            Caption         =   "&Above"
         End
         Begin VB.Menu insBelow 
            Caption         =   "&Below"
         End
      End
      Begin VB.Menu lvlDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu showChar 
         Caption         =   "&Characters"
      End
   End
   Begin VB.Menu properties 
      Caption         =   "&Properties"
      Begin VB.Menu lvlSchemes 
         Caption         =   "Set &Schemes"
      End
      Begin VB.Menu mnuPacProp 
         Caption         =   "&PacMan Properties"
      End
      Begin VB.Menu mnuGhostProp 
         Caption         =   "&Ghost Properties"
      End
      Begin VB.Menu mnuItemProp 
         Caption         =   "&Item Properties"
      End
      Begin VB.Menu mnuLevelProp 
         Caption         =   "&Level Properties"
      End
   End
   Begin VB.Menu aid 
      Caption         =   "&Aid"
      Begin VB.Menu NoClickAid 
         Caption         =   "&No-Click Aid"
         Shortcut        =   {F8}
      End
      Begin VB.Menu spread 
         Caption         =   "&Spreads"
         Begin VB.Menu spreadempty 
            Caption         =   "&Clear/Empty"
         End
         Begin VB.Menu spreadfood 
            Caption         =   "&Food"
         End
         Begin VB.Menu spreadshiel 
            Caption         =   "&Shield"
         End
         Begin VB.Menu spreadwall1 
            Caption         =   "&Wall 1"
         End
         Begin VB.Menu spreadwall2 
            Caption         =   "&Wall 2"
         End
      End
   End
End
Attribute VB_Name = "frmLevelEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Init
    InitDirectDraw
    lvledInitLevelView
    
    lvledLoadLevel 0
    Do
        Blt
        DoEvents
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub insAbove_Click()
    InsertLevel True
End Sub

Private Sub insBelow_Click()
    InsertLevel False
End Sub

Private Sub lvlDelete_Click()
    DeleteLevel
End Sub

Private Sub lvlSave_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Do you want to save on level " + LTrim$(Str$(SelectedLevel)) + "?", vbYesNo, "Save Confirmation")
    If result = vbYes Then lvledSaveLevel
End Sub

Private Sub lvlSchemes_Click()
    Load frmLevelEditSchemes
    frmLevelEditSchemes.Show
End Sub

Private Sub lvlSelect_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Do you want to load level " + LTrim$(Str$(SelectedLevel)) + " to surface?" + vbCrLf + "Your current work will be overwriten", vbYesNo, "Load Confirmation")
    If result = vbYes Then lvledLoadLevel SelectedLevel
End Sub

Private Sub lvlSurface_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicX = Int(X / 22)
    PicY = Int(Y / 22)
    lvledSetSurface PicX, PicY, curItem
End Sub

Private Sub lvlSurface_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicX = Int(X / 22)
    PicY = Int(Y / 22)
    If NoClickAid.Checked Then lvledSetSurface PicX, PicY, curItem
    StatusBar.Panels(1).Text = LTrim$(Str$(PicX)) + "," + LTrim$(Str$(PicY))
End Sub

Private Sub lvlView_NodeClick(ByVal Node As MSComctlLib.Node)
    mnuLevel.Enabled = True
    PopupMenu mnuLevel
End Sub

Private Sub mnuGhostProp_Click()
    Load GhostProp
    GhostProp.Show
End Sub

Private Sub mnuItemProp_Click()
    Load ItemProp
    ItemProp.Show
End Sub

Private Sub mnuLevelProp_Click()
    Load LevelProp
    LevelProp.Show
End Sub

Private Sub mnuPacProp_Click()
    Load PacProp
    PacProp.Show
End Sub

Private Sub NoClickAid_Click()
    If NoClickAid.Checked Then NoClickAid.Checked = False Else NoClickAid.Checked = True
End Sub

Private Sub showChar_Click()
    If showChar.Checked Then showChar.Checked = False Else showChar.Checked = True
End Sub

Private Sub spreadempty_Click()
    SpreadSurface 0
End Sub

Private Sub spreadfood_Click()
    SpreadSurface 1
End Sub

Private Sub spreadshiel_Click()
    SpreadSurface 2
End Sub

Private Sub spreadwall1_Click()
    SpreadSurface 3
End Sub

Private Sub spreadwall2_Click()
    SpreadSurface 4
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case Toolbar.buttons(1): curItem = 1    'Food Selected
        Case Toolbar.buttons(2): curItem = 2    'Shield Selected
        Case Toolbar.buttons(3): curItem = 3    'Wall 1 Selected
        Case Toolbar.buttons(4): curItem = 4    'Wall 2 Selected
        Case Toolbar.buttons(5): curItem = 0    'Erase Selected
    End Select
End Sub
