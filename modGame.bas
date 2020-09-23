Attribute VB_Name = "modGame"
Option Explicit

Sub InitGame()
    Pac.Life = PacLifeStart
    Pac.Score = 0
    Pac.ShieldTime = 0
    Pac.Level = PacLevelStart
End Sub

Sub InitLevel()
    LoadLevel Pac.Level
End Sub

Sub InitStartPos()
    Pac.x = Pac.StartX: Pac.y = Pac.StartY: Pac.Delay = Pac.StartDelay
    Pac.xDir = Pac.StartxDir: Pac.yDir = Pac.StartyDir
    Pac.DrunkTime = 0: Pac.ShieldTime = 0: Pac.MouthOpen = False
    Pac.NextDir = 0
    Dim GNo
    For GNo = 1 To 4
        Ghost(GNo).x = Ghost(GNo).StartX
        Ghost(GNo).y = Ghost(GNo).StartY
        Ghost(GNo).yDir = 0: Ghost(GNo).xDir = 0
        Ghost(GNo).Delay = Ghost(GNo).StartDelay
        Ghost(GNo).Sick = False
    Next GNo
End Sub

Sub MovOther()
Dim x, y, GhostNo
Erase aGame2()

With Game
    If .Beer.Appear Then aGame2(.Beer.x, .Beer.y) = pac_Beer
    If .Berry.Appear Then aGame2(.Berry.x, .Berry.y) = pac_Berry
    If .Cherry.Appear Then aGame2(.Cherry.x, .Cherry.y) = pac_Cherry
    If .Life.Appear Then aGame2(.Life.x, .Life.y) = pac_Life
End With

For GhostNo = 1 To 4
    If Ghost(GhostNo).Sick Then
        aGame2(Ghost(GhostNo).x, Ghost(GhostNo).y) = pac_Ghost(5, GhostNo)
    Else
        aGame2(Ghost(GhostNo).x, Ghost(GhostNo).y) = pac_Ghost(GhostNo, 1)
    End If
Next GhostNo
    
'====================================================================
Dim RandomX, RandomY, ItemNo
Dim Items As ItemChar
    For ItemNo = 1 To 4
        Select Case ItemNo
            Case 1: Items = Game.Beer
            Case 2: Items = Game.Berry
            Case 3: Items = Game.Cherry
            Case 4: Items = Game.Life
        End Select
        
        With Items
            .CurrentTime = .CurrentTime + 1
            If .Appear Then
                If .CurrentTime > .AppearTime + AdjustingSpeed Then
                    .CurrentTime = 0
                    .Appear = False
                End If
            Else
                If .Amount > 0 And .CurrentTime > .Delay + AdjustingSpeed Then
                    .Amount = .Amount - 1
                    Do
                        Randomize Timer
                        RandomX = Int(Rnd * MaxGameX)
                        RandomY = Int(Rnd * MaxGameY)
                    Loop While aGame(RandomX, RandomY) = pac_Wall Or aGame(RandomX, RandomY) = pac_Wall2 Or aGame(RandomX, RandomY) = pac_Shield
                    .x = RandomX
                    .y = RandomY
                    .Appear = True
                End If
            End If
        End With
        
        Select Case ItemNo
            Case 1: Game.Beer = Items
            Case 2: Game.Berry = Items
            Case 3: Game.Cherry = Items
            Case 4: Game.Life = Items
        End Select
        
    Next ItemNo
End Sub

Sub GameBody()
    If frmMain.PauseGame.Checked = True Then
        WhereAreWe = 2
        PauseAnimation
        WhereAreWe = 1
    End If
    MovOther
    MovePac
    MoveGhost
    CheckDeath
    GameBlt
    frmMain.StatusBar.Panels(1).Text = "LIFE : " + LTrim$(Str$(Pac.Life))
    frmMain.StatusBar.Panels(3).Text = Format(Pac.Score * 10, "###,###,###")
End Sub

