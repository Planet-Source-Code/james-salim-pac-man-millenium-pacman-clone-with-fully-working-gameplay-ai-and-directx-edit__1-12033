Attribute VB_Name = "modGameLevel"
Type LevelPacChar
    StartDelay As Integer
    DrunkDelay As Integer
    x As Byte
    y As Byte
    xDir As Integer
    yDir As Integer
    ProtectTime As Integer
    DrunkTime As Integer
End Type

Type LevelGhostChar
    x As Byte
    y As Byte
    Delay As Integer
    SickDelay As Integer
End Type

Type LevelItemChar
    AppearTime As Integer
    Delay As Integer
    Amount As Integer
End Type

Type LevelItems
    Beer As LevelItemChar
    Berry As LevelItemChar
    Cherry As LevelItemChar
    Life As LevelItemChar
End Type
 
Type LevelScheme
    Back As Byte
    Food As Byte
    Wall1 As Byte
    Wall2 As Byte
End Type

Public Type LevelBody
    LvlNo As Integer
    lvlName As String * 32
    lvlSurf(18, 18) As Byte
    lvlScheme As LevelScheme
    lvlPac As LevelPacChar
    lvlGhost(1 To 4) As LevelGhostChar
    lvlItems As LevelItems
End Type

Dim Body As LevelBody
Dim TempItem As ItemChar
Dim TempLevelItem As LevelItemChar
Dim Filenumber As Long

Function LoadLevel(LevelNo As Integer) As Boolean
    Filenumber = FreeFile
    
    On Error GoTo ErrorProc

    Open fpGame + "level.dat" For Random Access Read As #Filenumber Len = 461
        Get #Filenumber, 1, Body
        MaxLevelNo = Body.LvlNo
        
        If RTrim$(Body.lvlName) <> "PMLVL" Then GoTo ErrorProc
        If LevelNo > MaxLevelNo Or LevelNo <= 0 Then GoTo ErrorProc
        
        With Body
            Get #Filenumber, LevelNo + 1, Body
            Game.Point_on_Arena = 0
            For y = 0 To 18
                For x = 0 To 18
                    Select Case .lvlSurf(x, y)
                        Case 1
                            aGame(x, y) = pac_Food
                            Game.Point_on_Arena = Game.Point_on_Arena + 1
                        Case 2
                            aGame(x, y) = pac_Shield
                            Game.Point_on_Arena = Game.Point_on_Arena + 1
                        Case 3: aGame(x, y) = pac_Wall
                        Case 4: aGame(x, y) = pac_Wall2
                        Case Else: aGame(x, y) = 0
                    End Select
                Next x
            Next y
                
            InitLevelSurfaces .lvlScheme
    
            With .lvlPac
                Pac.StartDelay = .StartDelay
                Pac.DrunkDelay = .DrunkDelay
                Pac.StartX = .x: Pac.StartY = .y
                Pac.StartxDir = .xDir
                Pac.StartyDir = .yDir
                ProtectTime = .ProtectTime
                DrunkTime = .DrunkTime
            End With
    
            For GNo = 1 To 4
                Ghost(GNo).StartX = .lvlGhost(GNo).x
                Ghost(GNo).StartY = .lvlGhost(GNo).y
                Ghost(GNo).StartDelay = .lvlGhost(GNo).Delay
                Ghost(GNo).SickDelay = .lvlGhost(GNo).SickDelay
            Next GNo
                
            For ItemNo = 1 To 4
                Select Case ItemNo
                    Case 1: TempLevelItem = .lvlItems.Beer
                    Case 2: TempLevelItem = .lvlItems.Berry
                    Case 3: TempLevelItem = .lvlItems.Cherry
                    Case 4: TempLevelItem = .lvlItems.Life
                End Select
                        
                TempItem.AppearTime = TempLevelItem.AppearTime
                TempItem.Delay = TempLevelItem.Delay
                TempItem.Amount = TempLevelItem.Amount
                        
                Select Case ItemNo
                    Case 1: Game.Beer = TempItem
                    Case 2: Game.Berry = TempItem
                    Case 3: Game.Cherry = TempItem
                    Case 4: Game.Life = TempItem
                End Select
            Next ItemNo
        End With
    Close #Filenumber
     
    LoadLevel = True
    Exit Function
    
ErrorProc:

LoadLevel = False
End Function

