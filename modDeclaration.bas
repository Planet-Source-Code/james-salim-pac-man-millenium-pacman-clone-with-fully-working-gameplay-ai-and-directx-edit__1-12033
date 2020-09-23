Attribute VB_Name = "modDeclaration"

Public fpImage As String
Public fpGame As String

Public GameStart As Boolean
Public GameRestart As Boolean
'=====================================================================
'=====================================================================
'=GAME DECLARATION====================================================
'=====================================================================

Public Const MaxGameX = 18
Public Const MaxGameY = 18

Public NeedRepaint As Integer

Public aGame(MaxGameX, MaxGameY) As Integer     'Arena Layout without the characters
Public aGame2(MaxGameX, MaxGameY) As Integer    'Arena Layout only the characters and items
Public aBlt(MaxGameX, MaxGameY) As Integer      'Arena layout with all characters, etc this is
                                                'read by Blt() command, so it shall have all information
                                                'On what to show on the screen

Type ItemChar                   'Type for Items
    Amount As Integer
    Delay As Integer
    CurrentTime As Integer
    AppearTime As Integer
    
    x As Integer: y As Integer
    Appear As Boolean
End Type
                                
Type GameChar
    GhostEatCombo As Integer
    Point_on_Arena As Integer
    SpecialTime As Integer
    Berry As ItemChar
    Cherry As ItemChar
    Beer As ItemChar
    Life As ItemChar
End Type

Type PacChar
    Level As Integer
    Life As Integer
    Score As Long: ScoreToLife As Integer
    
    NextDir As Integer
    StartX As Integer:  StartY As Integer:  StartDelay As Integer
    StartxDir As Integer: StartyDir As Integer
    x As Integer:       y As Integer
    
    xDir As Integer:    yDir As Integer
    HeartBeat As Integer
    Delay As Integer
    DrunkTime As Integer: DrunkDelay As Integer
    ShieldTime As Integer
    MouthOpen As Boolean
    Dead As Boolean
End Type
    
Type GhostChar
    StartX As Integer: StartY As Integer: StartDelay As Integer
    x As Integer: y As Integer
    xDir As Integer: yDir As Integer
    LastX(10) As Integer: LastY(10) As Integer
    PacLastX As Integer: PacLastY As Integer
    HeartBeat As Integer
    Delay As Integer
    Sick As Boolean: SickDelay As Integer
End Type

Type Direction
    Possibility As Boolean
    Favour As Integer
    Percentage As Integer
    xDir As Integer
    yDir As Integer
End Type

Public Pac As PacChar
Public Ghost(1 To 4) As GhostChar
Public Game As GameChar
Public DrunkTime As Integer, ProtectTime As Integer
Public WhereAreWe As Integer

Public PacLifeStart As Integer
Public PacScoretoLife As Integer
Public PacLevelStart As Integer
Public PacDrunkDelay As Integer

Public Play_bPoint  As Boolean
Public Play_bDrunk  As Boolean
Public Play_bShield As Boolean
Public Play_bKill   As Boolean
Public Play_bWin    As Boolean
Public Play_bDead   As Boolean

Public MaxLevelNo As Integer
Public AdjustingSpeed As Integer
Public GhostAgressivity As Integer
Public ReturnName As String
Public ReplySent As Boolean
Public Marquee(14) As String
'WhereAreWe
'0=TitleScreen
'1=Game
'2=Pause
'3=End

'Wall Config number
'0      Nothing
'------------------------------------
'250    Wall
'251    Wall 2
'------------------------------------
'240    Point
'241    Special Point / Shield
'242    Item Berry
'243    Item Cherry
'244    Item Live
'245    Item Beer
'------------------------------------
'11-14  Ghost 1(Red)
'21-24  Ghost 2(Cyan)
'31-34  Ghost 3(Green)
'41-44  Ghost 4(Yellow)
'51     Ghost Sick (50 + GhostNo)
'101-104    Pac Man close mouth
'111-114    Pac Man open mouth
'121-124    Pac Man close mouth w/ shield
'131-134    Pac Man open mouth w/ shield
'------------------------------------

Public Const pac_Nothing = 0
Public Const pac_Wall = 250
Public Const pac_Wall2 = 251
Public Const pac_Food = 240
Public Const pac_Shield = 241
Public Const pac_Berry = 242
Public Const pac_Cherry = 243
Public Const pac_Life = 244
Public Const pac_Beer = 245
Public pac_Ghost(1 To 5, 1 To 4)
Public pac_Pac(1 To 2, 1 To 4)
                             
Sub InitDeclaration()
    Dim GhostNo, Direction
    For GhostNo = 1 To 5
        For Direction = 1 To 4
            pac_Ghost(GhostNo, Direction) = GhostNo * 10 + Direction
        Next Direction
    Next GhostNo
    
    For GhostNo = 1 To 2
        For Direction = 1 To 4
            pac_Pac(GhostNo, Direction) = ((GhostNo + 9) * 10) + Direction
        Next Direction
    Next GhostNo
    
    frmMain.Arena.Width = 418
    frmMain.Arena.Height = 418
    frmMain.ScaleWidth = 418
    frmMain.ScaleHeight = 418
    
    fpImage = App.Path + "\pacman\image\"
    fpGame = App.Path + "\pacman\game\"
    
    PacLifeStart = 5
    PacScoretoLife = 1000
    PacLevelStart = 1
    AdjustingSpeed = 0
    MaxLevelNo = 10
    GhostAgressivity = 5
    
    GameRestart = False
End Sub
