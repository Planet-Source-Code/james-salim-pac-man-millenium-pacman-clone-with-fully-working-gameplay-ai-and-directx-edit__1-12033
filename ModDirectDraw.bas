Attribute VB_Name = "ModDirectDraw"
Option Explicit
Dim SurfDesc1 As DDSURFACEDESC2
Dim SurfDesc2 As DDSURFACEDESC2
Dim SurfDesc3 As DDSURFACEDESC2

Dim Key As DDCOLORKEY

'Sub-program to be loaded by InitDirectX()
Sub InitDirectDraw()
    Set DDraw = DirectX.DirectDrawCreate("")
    frmMain.Show
    DDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
         
    '------Init Primary Surface------
    With SurfDesc1
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    Set sPrimary = DDraw.CreateSurface(SurfDesc1)
    Set ddClipper = DDraw.CreateClipper(0)
    ddClipper.SetHWnd frmMain.Arena.hWnd
    sPrimary.SetClipper ddClipper
    sPrimary.SetForeColor QBColor(15)
    
    With SurfDesc1
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 418
        .lWidth = 418
    End With
    Set sBack = DDraw.CreateSurface(SurfDesc1)
    Set sTemp = DDraw.CreateSurface(SurfDesc1)
    sBack.SetForeColor QBColor(15)
    
    InitGameSurfaces
    InitNonGameSurfaces
End Sub

Sub InitGameSurfaces()
    '------Init Sprite Surface------
    '======Set Surface Description=====
    With SurfDesc2
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 22
        .lWidth = 22
    End With
    
    '=====Load Bitmap File to surface======
    With DDraw
        'Load PacMan Character to Surface Memory
        Set sPac(1, 1) = .CreateSurfaceFromFile(fpImage + "pac\close_up.bmp", SurfDesc2)
        Set sPac(1, 2) = .CreateSurfaceFromFile(fpImage + "pac\close_dn.bmp", SurfDesc2)
        Set sPac(1, 3) = .CreateSurfaceFromFile(fpImage + "pac\close_lf.bmp", SurfDesc2)
        Set sPac(1, 4) = .CreateSurfaceFromFile(fpImage + "pac\close_rg.bmp", SurfDesc2)
        Set sPac(2, 1) = .CreateSurfaceFromFile(fpImage + "pac\open_up.bmp", SurfDesc2)
        Set sPac(2, 2) = .CreateSurfaceFromFile(fpImage + "pac\open_dn.bmp ", SurfDesc2)
        Set sPac(2, 3) = .CreateSurfaceFromFile(fpImage + "pac\open_lf.bmp", SurfDesc2)
        Set sPac(2, 4) = .CreateSurfaceFromFile(fpImage + "pac\open_rg.bmp", SurfDesc2)
        Set sPacDead = .CreateSurfaceFromFile(fpImage + "pac\skull.bmp", SurfDesc2)
        
        'Load Ghost Character to surface Memory
        Set sGhost(1, 1) = .CreateSurfaceFromFile(fpImage + "ghost\gred_up.bmp", SurfDesc2)
        Set sGhost(1, 2) = .CreateSurfaceFromFile(fpImage + "ghost\gred_dn.bmp", SurfDesc2)
        Set sGhost(1, 3) = .CreateSurfaceFromFile(fpImage + "ghost\gred_lf.bmp", SurfDesc2)
        Set sGhost(1, 4) = .CreateSurfaceFromFile(fpImage + "ghost\gred_rg.bmp", SurfDesc2)
        
        Set sGhost(2, 1) = .CreateSurfaceFromFile(fpImage + "ghost\gcyan_up.bmp", SurfDesc2)
        Set sGhost(2, 2) = .CreateSurfaceFromFile(fpImage + "ghost\gcyan_dn.bmp", SurfDesc2)
        Set sGhost(2, 3) = .CreateSurfaceFromFile(fpImage + "ghost\gcyan_lf.bmp", SurfDesc2)
        Set sGhost(2, 4) = .CreateSurfaceFromFile(fpImage + "ghost\gcyan_rg.bmp", SurfDesc2)
        
        Set sGhost(3, 1) = .CreateSurfaceFromFile(fpImage + "ghost\ggreen_up.bmp", SurfDesc2)
        Set sGhost(3, 2) = .CreateSurfaceFromFile(fpImage + "ghost\ggreen_dn.bmp", SurfDesc2)
        Set sGhost(3, 3) = .CreateSurfaceFromFile(fpImage + "ghost\ggreen_lf.bmp", SurfDesc2)
        Set sGhost(3, 4) = .CreateSurfaceFromFile(fpImage + "ghost\ggreen_rg.bmp", SurfDesc2)
       
        Set sGhost(4, 1) = .CreateSurfaceFromFile(fpImage + "ghost\gyellow_up.bmp", SurfDesc2)
        Set sGhost(4, 2) = .CreateSurfaceFromFile(fpImage + "ghost\gyellow_dn.bmp", SurfDesc2)
        Set sGhost(4, 3) = .CreateSurfaceFromFile(fpImage + "ghost\gyellow_lf.bmp", SurfDesc2)
        Set sGhost(4, 4) = .CreateSurfaceFromFile(fpImage + "ghost\gyellow_rg.bmp", SurfDesc2)
       
        Set sGhost(5, 1) = .CreateSurfaceFromFile(fpImage + "ghost\gsick.bmp", SurfDesc2)
        
        Set sItemShield = .CreateSurfaceFromFile(fpImage + "item\protect.bmp", SurfDesc2)
        Set sItemLife = .CreateSurfaceFromFile(fpImage + "item\1up.bmp", SurfDesc2)
        Set sItemBerry = .CreateSurfaceFromFile(fpImage + "item\berry.bmp", SurfDesc2)
        Set sItemCherry = .CreateSurfaceFromFile(fpImage + "item\cherry.bmp", SurfDesc2)
        Set sItemBeer = .CreateSurfaceFromFile(fpImage + "item\beer.bmp", SurfDesc2)
        Set sPacShield = .CreateSurfaceFromFile(fpImage + "item\shield.bmp", SurfDesc2)
    End With
    
    '=======Set Color Key for Sprites=======
    Key.low = RGB(255, 255, 255)
    Key.high = RGB(255, 255, 255)
           
    Dim a, b
    For a = 1 To 2
        For b = 1 To 4
            sPac(a, b).SetColorKey DDCKEY_SRCBLT, Key
        Next b
    Next a
    For a = 1 To 4
        For b = 1 To 4
            sGhost(a, b).SetColorKey DDCKEY_SRCBLT, Key
        Next b
    Next a
    sGhost(5, 1).SetColorKey DDCKEY_SRCBLT, Key
    sItemShield.SetColorKey DDCKEY_SRCBLT, Key
    sItemLife.SetColorKey DDCKEY_SRCBLT, Key
    sItemBerry.SetColorKey DDCKEY_SRCBLT, Key
    sItemCherry.SetColorKey DDCKEY_SRCBLT, Key
    sItemBeer.SetColorKey DDCKEY_SRCBLT, Key
    sPacShield.SetColorKey DDCKEY_SRCBLT, Key
End Sub

Sub InitNonGameSurfaces()
    With SurfDesc3
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 418
        .lWidth = 418
    End With
    Set sTitleScreen = DDraw.CreateSurfaceFromFile(fpImage + "title.bmp", SurfDesc3)
    
    SurfDesc3.lWidth = 200
    SurfDesc3.lHeight = 60
    Set sPause = DDraw.CreateSurfaceFromFile(fpImage + "pause.bmp", SurfDesc3)
    
    '=======Set Color Key for Sprites=======
    'Key.low = RGB(255, 255, 255)
           
    'sPause.SetColorKey DDCKEY_SRCBLT, Key

End Sub

Sub InitLevelSurfaces(schm As LevelScheme)

    Set sFood = Nothing
    Set sWall = Nothing
    Set sWall2 = Nothing
    Set sBackground = Nothing
    
    With SurfDesc1
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 22
        .lWidth = 22
    End With
    With SurfDesc2
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 418
        .lWidth = 418
    End With
    
    Set sFood = DDraw.CreateSurfaceFromFile(fpImage + "schemes\" + LTrim$(Str$(schm.Food)) + "_food.bmp", SurfDesc1)
    Set sWall = DDraw.CreateSurfaceFromFile(fpImage + "schemes\" + LTrim$(Str$(schm.Wall1)) + "_wall.bmp", SurfDesc1)
    Set sWall2 = DDraw.CreateSurfaceFromFile(fpImage + "schemes\" + LTrim$(Str$(schm.Wall2)) + "_wall2.bmp", SurfDesc1)
    Set sBackground = DDraw.CreateSurfaceFromFile(fpImage + "schemes\" + LTrim$(Str$(schm.Back)) + "_Back.bmp", SurfDesc2)

    Key.low = RGB(255, 255, 255)
    Key.high = RGB(255, 255, 255)

    sFood.SetColorKey DDCKEY_SRCBLT, Key
    sWall.SetColorKey DDCKEY_SRCBLT, Key
    sWall2.SetColorKey DDCKEY_SRCBLT, Key
End Sub

Sub PreBlt()

Dim x, y, GhostNo, tDirection
For x = 0 To 18
    For y = 0 To 18
        aBlt(x, y) = aGame(x, y)
    Next y
Next x

If Pac.yDir = -1 Then tDirection = 1
If Pac.yDir = 1 Then tDirection = 2
If Pac.xDir = -1 Then tDirection = 3
If Pac.xDir = 1 Then tDirection = 4

If Pac.MouthOpen Then
    aBlt(Pac.x, Pac.y) = pac_Pac(2, tDirection)
Else
    aBlt(Pac.x, Pac.y) = pac_Pac(1, tDirection)
End If

If Game.Beer.Appear Then aBlt(Game.Beer.x, Game.Beer.y) = pac_Beer
If Game.Berry.Appear Then aBlt(Game.Berry.x, Game.Berry.y) = pac_Berry
If Game.Cherry.Appear Then aBlt(Game.Cherry.x, Game.Cherry.y) = pac_Cherry
If Game.Life.Appear Then aBlt(Game.Life.x, Game.Life.y) = pac_Life

For GhostNo = 1 To 4
    If Ghost(GhostNo).Sick Then
        If Pac.ShieldTime > 20 Then
            aBlt(Ghost(GhostNo).x, Ghost(GhostNo).y) = pac_Ghost(5, GhostNo)
        Else
            If Pac.ShieldTime / 2 = Int(Pac.ShieldTime / 2) Then
                If Ghost(GhostNo).yDir = 1 Then tDirection = 1 Else tDirection = 2
                If Ghost(GhostNo).xDir = 1 Then tDirection = 4 Else tDirection = 3
                aBlt(Ghost(GhostNo).x, Ghost(GhostNo).y) = pac_Ghost(GhostNo, tDirection)
            Else
                aBlt(Ghost(GhostNo).x, Ghost(GhostNo).y) = pac_Ghost(5, GhostNo)
            End If
        End If
    Else
        If Ghost(GhostNo).yDir = 1 Then tDirection = 1 Else tDirection = 2
        If Ghost(GhostNo).xDir = 1 Then tDirection = 4 Else tDirection = 3
        aBlt(Ghost(GhostNo).x, Ghost(GhostNo).y) = pac_Ghost(GhostNo, tDirection)
    End If
Next GhostNo
End Sub

Sub GameBlt()
    PreBlt
    '------Get Arena Location (RECT)------
    DirectX.GetWindowRect frmMain.Arena.hWnd, aRect
    aTop = aRect.Top
    aLeft = aRect.Left

    Dim rBack As RECT, rSprite As RECT, rTarget As RECT
    Dim x As Integer, y As Integer, GhostNo As Integer, ChrDir As Integer
    Dim sTemp As DirectDrawSurface7
    
    rBack.Top = 0
    rBack.Left = 0
    rBack.Bottom = 418
    rBack.Right = 418
    sBack.Blt rBack, sBackground, rBack, DDBLT_WAIT
    
    rSprite.Bottom = 22: rSprite.Right = 22
        
    For y = 0 To 18
        For x = 0 To 18
            Set sTemp = Nothing
            Select Case aBlt(x, y)
                Case pac_Wall: Set sTemp = sWall
                Case pac_Wall2: Set sTemp = sWall2
                Case pac_Food: Set sTemp = sFood
                Case pac_Shield: Set sTemp = sItemShield
                Case pac_Berry: Set sTemp = sItemBerry
                Case pac_Cherry: Set sTemp = sItemCherry
                Case pac_Life: Set sTemp = sItemLife
                Case pac_Beer: Set sTemp = sItemBeer
                Case pac_Ghost(1, 1) To pac_Ghost(4, 4)
                    GhostNo = Int(aBlt(x, y) / 10)
                    ChrDir = aBlt(x, y) - (GhostNo * 10)
                    Set sTemp = sGhost(GhostNo, ChrDir)
                Case pac_Ghost(5, 1) To pac_Ghost(5, 4): Set sTemp = sGhost(5, 1)
                Case pac_Pac(1, 1) To pac_Pac(2, 4)
                    If aBlt(x, y) > 110 Then
                        Set sTemp = sPac(2, aBlt(x, y) - 110)
                    Else
                        Set sTemp = sPac(1, aBlt(x, y) - 100)
                    End If
                Case Else: GoTo SkipFor
            End Select
            
            rTarget.Top = y * 22
            rTarget.Left = x * 22
            rTarget.Bottom = (y + 1) * 22
            rTarget.Right = (x + 1) * 22
            sBack.Blt rTarget, sTemp, rSprite, DDBLT_KEYSRC Or DDBLT_WAIT
SkipFor:
        Next x
    Next y
    
    sPrimary.Blt aRect, sBack, rBack, DDBLT_WAIT
End Sub

Sub PauseAnimation()
        Dim Anim_x, Anim_y, Anim_xDir, Anim_yDir, Moment
        Dim rPause As RECT, rTarget As RECT, rTemp As RECT
        Anim_x = 109: Anim_y = 179: Anim_xDir = 1: Anim_yDir = 1
        rPause.Bottom = 60: rPause.Right = 200
        rTemp.Bottom = 418: rTemp.Right = 418
        Moment = 0

        Do
            Moment = Moment + 1
            If Moment > 5000 Then
                Moment = 0
                Anim_x = Anim_x + Anim_xDir
                Anim_y = Anim_y + Anim_yDir
                If Anim_x + 200 > 417 Then Anim_xDir = Anim_xDir * -1
                If Anim_y + 60 > 417 Then Anim_yDir = Anim_yDir * -1
                If Anim_x < 1 Then Anim_xDir = Anim_xDir * -1
                If Anim_y < 1 Then Anim_yDir = Anim_yDir * -1
        
                rTarget.Top = Anim_y: rTarget.Left = Anim_x
                rTarget.Bottom = Anim_y + 60: rTarget.Right = Anim_x + 200
            
                sTemp.Blt rTemp, sBack, rTemp, DDBLT_WAIT
                sTemp.Blt rTarget, sPause, rPause, DDBLT_WAIT
                
                DirectX.GetWindowRect frmMain.Arena.hWnd, aRect
                sPrimary.Blt aRect, sTemp, rTemp, DDBLT_WAIT
                DoEvents
            End If
        Loop Until Not (frmMain.PauseGame.Checked)
End Sub
