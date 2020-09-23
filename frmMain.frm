VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PacMan Millenium"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6255
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   6255
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "LIFE : "
            TextSave        =   "LIFE : "
            Object.ToolTipText     =   "This show Pac-man's Current Life"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   5768
            MinWidth        =   5768
            Text            =   "Pac Man Millenium"
            TextSave        =   "Pac Man Millenium"
            Object.ToolTipText     =   "Poem by First Church of Pac Man"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "This show Pac-man's Current Score"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Arena 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6270
      Left            =   0
      ScaleHeight     =   418
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   418
      TabIndex        =   0
      Top             =   0
      Width           =   6270
      Begin MSComDlg.CommonDialog DlgHelp 
         Left            =   5640
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lFont 
         BackStyle       =   0  'Transparent
         Caption         =   "FontSample"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   5520
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Menu GameMenu 
      Caption         =   "&Game"
      Begin VB.Menu NewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu RestartGame 
         Caption         =   "&Restart Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu PauseGame 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu Seperator1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Option 
      Caption         =   "&Option"
      Begin VB.Menu Sound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu Music 
         Caption         =   "&Music"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu view 
         Caption         =   "&View"
         Begin VB.Menu view100 
            Caption         =   "&Normal"
            Shortcut        =   {F5}
         End
         Begin VB.Menu view150 
            Caption         =   "150%"
            Shortcut        =   {F6}
         End
         Begin VB.Menu view200 
            Caption         =   "200%"
            Shortcut        =   {F7}
         End
      End
      Begin VB.Menu Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu Game_Configuration 
         Caption         =   "&Game Configuration"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu Content 
         Caption         =   "&Content"
         Shortcut        =   {F1}
      End
      Begin VB.Menu HelponHelp 
         Caption         =   "&Help on help"
      End
      Begin VB.Menu Seperator4 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About PacMan Millenium"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===Load the about form===
Private Sub About_Click()
    Load frmAbout
    frmAbout.Show
    PauseGame.Checked = True
End Sub

'===Trigerred upon key pressed===
'Initiate movement of pacman
Private Sub Arena_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: Pac.NextDir = 1     'Up Arrow Pressed
        Case 40: Pac.NextDir = 2     'Down Arrow Pressed
        Case 37: Pac.NextDir = 3     'Left Arrow Pressed
        Case 39: Pac.NextDir = 4     'Right Arrow Pressed
    End Select
End Sub

Private Sub Content_Click()
App.HelpFile = App.Path + "pacman\pacman.hlp"
DlgHelp.HelpCommand = 3
DlgHelp.ShowHelp
End Sub

Private Sub Exit_Click()
    Dim Result As VbMsgBoxResult
    Result = MsgBox("Do you really want to quit Pac-Man millenium?", vbYesNo, "Exit Pac Man Millenium")
    If Result = vbYes Then Unload frmMain
End Sub

Private Sub Form_Load()
DlgHelp.HelpFile = App.HelpFile
GameStart = False
InitDeclaration
InitDirectX
SetViewMode 0
    
DoEvents
'=============================END LOW LEVEL INIT=========

'=TITLE SCREEN SHOW=
    DirectX.GetWindowRect frmMain.Arena.hWnd, aRect
    Dim rTitle As RECT
    
    rTitle.Bottom = 418
    rTitle.Right = 418
    
Do
    If NeedRepaint = 1 Then sPrimary.Blt aRect, sTitleScreen, rTitle, DDBLT_WAIT: NeedRepaint = 0
    DoEvents
Loop Until GameStart

'=THE GAME IT SELF=
GameStart:
            NewGame.Enabled = False
            PauseGame.Enabled = True
            RestartGame.Enabled = True
            WhereAreWe = 1
            GameStart = False
            InitGame
            
GameRestart:    InitLevel
LevelReStart:   InitStartPos
                GameBlt
                PauseGame.Checked = True
                'Draw somethign to show that game paused,
                'eg. "Press F3 to continue
            
Do
    GameBody
    DoEvents
Loop Until Game.Point_on_Arena <= 0 Or Pac.Dead Or GameRestart

'=CHECK IF RESTART GAME IS REQUESTED=
If GameRestart Then GameRestart = False: GoTo GameStart

'=PROCEDURE IF PAC IS DEAD=
If Pac.Dead Then
    Pac.Dead = False
    Pac.Life = Pac.Life - 1
    If Pac.Life >= 0 Then GoTo LevelReStart
Else
    Pac.Level = Pac.Level + 1
    If Pac.Level <= MaxLevelNo Then GoTo GameRestart
    Pac.Level = Pac.Level - 1
End If


'=ENDING CREDIT/HIGHSCORE SCREEN=
GameOver:
    LoadHScore
    If Pac.Score > HScoreMinValue Then
        ReturnName = ""
        ReplySent = False
        Load HScoreAsk
        HScoreAsk.Show
        Do
            DoEvents
        Loop Until ReplySent
        SaveHScore ReturnName, Pac.Level, Pac.Score
    End If
    LoadHScore
    
    RestartGame.Enabled = False
    PauseGame.Enabled = False
    NewGame.Enabled = True
    WhereAreWe = 3
        
    Do
        rTitle.Bottom = 418
        rTitle.Right = 418
        sTemp.SetFontTransparency True
        sTemp.SetFont lFont.Font
        sTemp.SetForeColor RGB(255, 255, 255)
        
        sTemp.BltColorFill rTitle, QBColor(0)
        sTemp.DrawText 130, 10, "Hall-of-Fame", False
        
        For i = 1 To 10
            sTemp.DrawText 30, 40 + i * 32, HScoreList(i).Name, False
            sTemp.DrawText 270, 40 + i * 32, HScoreList(i).LastLevel, False
            sTemp.DrawText 320, 40 + i * 32, Format(HScoreList(i).Score * 10, "###,###,###"), False
        Next i
        
        DirectX.GetWindowRect frmMain.Arena.hWnd, aRect
        sPrimary.Blt aRect, sTemp, rTitle, DDBLT_WAIT
        DoEvents
    Loop Until GameStart
    
    GoTo GameStart
    
End Sub

Private Sub Form_Paint()
    NeedRepaint = 1
End Sub

Private Sub Form_Resize()
    StatusBar.Panels(2).Width = Me.ScaleWidth - 200
End Sub

Private Sub Game_Configuration_Click()
    PauseGame.Checked = True
    Load frmGameConfig
    frmGameConfig.Show
End Sub

Private Sub HelponHelp_Click()
    DlgHelp.HelpCommand = 4
    DlgHelp.ShowHelp
End Sub

Private Sub NewGame_Click()
    GameStart = True
End Sub

Private Sub PauseGame_Click()
    If PauseGame.Checked Then PauseGame.Checked = False Else PauseGame.Checked = True
End Sub

Private Sub RestartGame_Click()
    Dim Result As VbMsgBoxResult
    Result = MsgBox("A Game is currently in progress. Are you sure you want to restart the game?", vbYesNo, "Restart Game")
    If Result = vbYes Then
        GameRestart = True
    End If
End Sub

Private Sub view100_Click()
    SetViewMode 0
End Sub

Private Sub view150_Click()
    SetViewMode 1
End Sub

Private Sub view200_Click()
    SetViewMode 2
End Sub
