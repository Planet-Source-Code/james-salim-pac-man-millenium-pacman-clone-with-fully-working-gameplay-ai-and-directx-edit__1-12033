Attribute VB_Name = "ModGameChar"
Option Explicit

Sub CheckDeath()
    Dim GhostInPos, GhostNo
    GhostInPos = aGame2(Pac.x, Pac.y)
    Select Case GhostInPos
        Case pac_Ghost(1, 1) To pac_Ghost(4, 4)
            Play_bDead = True
            Pac.Dead = True
            Game.GhostEatCombo = 0
        Case pac_Ghost(5, 1) To pac_Ghost(5, 4)
            Play_bKill = True
            GhostNo = GhostInPos - 50
            Ghost(GhostNo).x = Ghost(GhostNo).StartX
            Ghost(GhostNo).y = Ghost(GhostNo).StartY
            Ghost(GhostNo).Sick = False
            Ghost(GhostNo).Delay = Ghost(GhostNo).StartDelay
            Game.GhostEatCombo = Game.GhostEatCombo + 1
            Pac.Score = Pac.Score + Game.GhostEatCombo * 20
            Pac.ScoreToLife = Pac.ScoreToLife + Game.GhostEatCombo * 20
    End Select
End Sub

Sub MovePac()
    '==================================================
    Pac.HeartBeat = Pac.HeartBeat + 1
    If Pac.HeartBeat >= (Pac.Delay + AdjustingSpeed) / 2 Then
        Pac.HeartBeat = 0
    Else
        Exit Sub
    End If
    If Pac.MouthOpen = True Then
        Pac.MouthOpen = False
        Exit Sub
    Else
        Pac.MouthOpen = True
    End If
    
    '===================================================
    If Pac.ScoreToLife > PacScoretoLife Then
        Play_bKill = True
        Pac.Life = Pac.Life + 1
        Pac.ScoreToLife = 0
    End If
    '===================================================
    Dim GhostNo
    If Pac.ShieldTime > 0 Then
        Pac.ShieldTime = Pac.ShieldTime - 1
        If Pac.ShieldTime = 0 Then
            For GhostNo = 1 To 4
                Ghost(GhostNo).Sick = False
                Ghost(GhostNo).Delay = Ghost(GhostNo).StartDelay
            Next GhostNo
            Game.GhostEatCombo = 0
        End If
    End If
        
    If Pac.DrunkTime > 0 Then
        Pac.DrunkTime = Pac.DrunkTime - 1
        If Pac.DrunkTime = 0 Then
            Pac.Delay = Pac.StartDelay
        End If
    End If
    '===================================================
    Dim NewX, NewY
    Dim NewXDir, NewYDir

    Select Case Pac.NextDir
        Case 1
            NewXDir = 0
            NewYDir = -1
        Case 2
            NewXDir = 0
            NewYDir = 1
        Case 3
            NewXDir = -1
            NewYDir = 0
        Case 4
            NewXDir = 1
            NewYDir = 0
        Case Else
            NewXDir = Pac.xDir
            NewYDir = Pac.yDir
    End Select
    
        NewX = Pac.x + NewXDir
        NewY = Pac.y + NewYDir
        If NewX < 0 Then NewX = MaxGameX
        If NewY < 0 Then NewY = MaxGameY
        If NewX > MaxGameX Then NewX = 0
        If NewY > MaxGameY Then NewY = 0
        If Not (aGame(NewX, NewY) = pac_Wall Or aGame(NewX, NewY) = pac_Wall2) Then
            Pac.xDir = NewXDir
            Pac.yDir = NewYDir
            Pac.x = NewX
            Pac.y = NewY
        Else
            NewX = Pac.x + Pac.xDir
            NewY = Pac.y + Pac.yDir
            If NewX < 0 Then NewX = MaxGameX
            If NewY < 0 Then NewY = MaxGameY
            If NewX > MaxGameX Then NewX = 0
            If NewY > MaxGameY Then NewY = 0
            
            If aGame(NewX, NewY) = pac_Wall Or aGame(NewX, NewY) = pac_Wall2 Then
                Pac.x = Pac.x: Pac.y = Pac.y
            Else
                Pac.x = NewX
                Pac.y = NewY
            End If
        End If
    
     
    '===================================================
    Dim ItemInPos
    ItemInPos = aGame(Pac.x, Pac.y)
    Select Case ItemInPos
        Case pac_Food
            Play_bPoint = True
            aGame(Pac.x, Pac.y) = pac_Nothing
            Game.Point_on_Arena = Game.Point_on_Arena - 1
            Pac.Score = Pac.Score + 1
            Pac.ScoreToLife = Pac.ScoreToLife + 1
        Case pac_Shield
            Play_bShield = True
            aGame(Pac.x, Pac.y) = pac_Nothing
            Game.Point_on_Arena = Game.Point_on_Arena - 1
            Pac.ShieldTime = ProtectTime
            For GhostNo = 1 To 4
                Ghost(GhostNo).Sick = True
                Ghost(GhostNo).Delay = Ghost(GhostNo).SickDelay
            Next GhostNo
            Pac.Score = Pac.Score + 5
            Pac.ScoreToLife = Pac.ScoreToLife + 5
    End Select
    
    ItemInPos = aGame2(Pac.x, Pac.y)
    Select Case ItemInPos
        Case pac_Beer
            Play_bDrunk = True
            Pac.DrunkTime = DrunkTime
            Pac.Delay = Pac.DrunkDelay
            Game.Beer.Appear = False
            Game.Beer.CurrentTime = 0
        Case pac_Berry
            Play_bWin = True
            Pac.Score = Pac.Score + 20
            Pac.ScoreToLife = Pac.Score + 20

            Game.Berry.Appear = False
            Game.Berry.CurrentTime = 0
        Case pac_Cherry
            Play_bWin = True
            Pac.Score = Pac.Score + 40
            Pac.ScoreToLife = Pac.ScoreToLife + 40
            Game.Cherry.Appear = False
            Game.Cherry.CurrentTime = 0
        Case pac_Life
            Play_bKill = True
            Pac.Life = Pac.Life + 1
            Game.Life.Appear = False
            Game.Life.CurrentTime = 0
    End Select
End Sub

Sub MoveGhost()
Attribute MoveGhost.VB_Description = "jjj"
    Dim GNo, NewX, NewY, x, y, DirNo, DirNoOpp(4), Posibility
    Dim MaxValue, MaxPercentage, ResultDir, Temp1, Temp2
    Dim Dir(4) As Direction
    
    DirNoOpp(1) = 2: DirNoOpp(2) = 1: DirNoOpp(3) = 4: DirNoOpp(4) = 3
    Dir(1).xDir = 0: Dir(2).xDir = 0: Dir(3).xDir = -1: Dir(4).xDir = 1
    Dir(1).yDir = -1: Dir(2).yDir = 1: Dir(3).yDir = 0: Dir(4).yDir = 0
    For GNo = 1 To 4
        With Ghost(GNo)
            .HeartBeat = .HeartBeat + 1
            If .HeartBeat >= .Delay + AdjustingSpeed Then
                .HeartBeat = 0
            Else
                GoTo NextGhost
            End If
            
            
            Dir(1).Possibility = True
            Dir(2).Possibility = True
            Dir(3).Possibility = True
            Dir(4).Possibility = True
            Dir(1).Percentage = 0
            Dir(2).Percentage = 0
            Dir(3).Percentage = 0
            Dir(4).Percentage = 0
            Dir(1).Favour = 0: Dir(2).Favour = 0
            Dir(3).Favour = 0: Dir(4).Favour = 0
            Posibility = 4
            
            If CheckWall(.x, .y - 1) Then Dir(1).Possibility = False
            If CheckWall(.x, .y + 1) Then Dir(2).Possibility = False
            If CheckWall(.x - 1, .y) Then Dir(3).Possibility = False
            If CheckWall(.x + 1, .y) Then Dir(4).Possibility = False

            For DirNo = 1 To 4
                If Dir(DirNo).Possibility Then
                    If .yDir = Dir(DirNo).yDir And .xDir = Dir(DirNo).xDir Then Dir(DirNoOpp(DirNo)).Favour = Dir(DirNoOpp(DirNo)).Favour - 1
                    Dir(DirNo).Favour = Dir(DirNo).Favour - GhostPast(GNo, .x + Dir(DirNo).xDir, .y + Dir(DirNo).yDir) * 2
                    y = .y: x = .x
                    If y = .PacLastY And x = .PacLastX Then .PacLastX = 0: .PacLastY = 0
                    Do
                        y = y + Dir(DirNo).yDir
                        x = x + Dir(DirNo).xDir
                        If x < 0 Or y < 0 Then Exit Do
                        If x > MaxGameX Or y > MaxGameY Then Exit Do
                        If CheckWall(x, y) Then Exit Do
                        If Pac.x = x And Pac.y = y Then
                            If .Sick Then
                                Dir(DirNo).Favour = Dir(DirNo).Favour - GhostAgressivity
                            Else
                                Dir(DirNo).Favour = Dir(DirNo).Favour + GhostAgressivity
                                .PacLastX = x: .PacLastY = y
                            End If
                        End If
                        If x = .PacLastX And y = .PacLastY Then Dir(DirNo).Favour = Dir(DirNo).Favour + 3
                        If aGame2(x, y) >= pac_Ghost(1, 1) And aGame2(x, y) <= pac_Ghost(5, 4) Then Dir(DirNo).Favour = Dir(DirNo).Favour - 1

                    Loop
                Else
                    Posibility = Posibility - 1
                    Dir(DirNo).Favour = -999
                End If
            Next DirNo
            If Posibility = 0 Then GoTo NextGhost
            
            MaxValue = GetMaxValue(Dir(1).Favour, Dir(2).Favour, Dir(3).Favour, Dir(4).Favour)
            
RandomFavour:
            For DirNo = 1 To 4
                If Dir(DirNo).Favour = MaxValue Then
                    Randomize Timer
                    Dir(DirNo).Percentage = Int(Rnd * 100)
                End If
            Next DirNo
            MaxPercentage = GetMaxValue(Dir(1).Percentage, Dir(2).Percentage, Dir(3).Percentage, Dir(4).Percentage)
            Temp1 = GetAmountEqualTo(Dir(1).Percentage, Dir(2).Percentage, Dir(3).Percentage, Dir(4).Percentage, MaxPercentage)
            If Temp1 > 1 Then GoTo RandomFavour
            
            For DirNo = 1 To 4
                If Dir(DirNo).Percentage = MaxPercentage Then ResultDir = DirNo: Exit For
            Next DirNo
            
            .xDir = Dir(ResultDir).xDir
            .yDir = Dir(ResultDir).yDir
            
            NewX = .x + .xDir
            NewY = .y + .yDir
            If NewX < 0 Then NewX = MaxGameX
            If NewY < 0 Then NewY = MaxGameY
            If NewX > MaxGameX Then NewX = 0
            If NewY > MaxGameY Then NewY = 0
                    
            .x = NewX
            .y = NewY
            GhostAddMove GNo, .x, .y
NextGhost:
        End With
    Next GNo
    
End Sub
