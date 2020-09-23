Attribute VB_Name = "ModDirectDraw"
Option Explicit
Public DirectX As New DirectX7
Public DDraw As DirectDraw7

Public sTemp As DirectDrawSurface7
Public sPrimary As DirectDrawSurface7
Public sBack  As DirectDrawSurface7
Public ddClipper As DirectDrawClipper

Public sBackground As DirectDrawSurface7    'Level schemes background surface
Public sFood As DirectDrawSurface7          'Level schemes food surface
Public sShield As DirectDrawSurface7        'Shield Item
Public sWall As DirectDrawSurface7          'Level schemes wall surface
Public sWall2 As DirectDrawSurface7         'Level schemes 2nd wall surface
Public sPac(1 To 4) As DirectDrawSurface7
Public sGhost(1 To 4) As DirectDrawSurface7
    
Dim SurfDesc1 As DDSURFACEDESC2
Dim SurfDesc2 As DDSURFACEDESC2
Dim SurfDesc3 As DDSURFACEDESC2

Dim Key As DDCOLORKEY

'Sub-program to be loaded by InitDirectX()
Sub InitDirectDraw()
    Set DDraw = DirectX.DirectDrawCreate("")
    frmLevelEditor.Show
    DDraw.SetCooperativeLevel frmLevelEditor.hWnd, DDSCL_NORMAL
         
    '------Init Primary Surface------
    With SurfDesc1
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    Set sPrimary = DDraw.CreateSurface(SurfDesc1)
    Set ddClipper = DDraw.CreateClipper(0)
    ddClipper.SetHWnd frmLevelEditor.lvlSurface.hWnd
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
    InitSurfaces
End Sub

Sub InitSurfaces()
    With SurfDesc1
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 22
        .lWidth = 22
    End With
    Set sShield = DDraw.CreateSurfaceFromFile(fpImage + "item\Protect.bmp", SurfDesc1)
    Set sPac(1) = DDraw.CreateSurfaceFromFile(fpImage + "pac\close_up.bmp", SurfDesc1)
    Set sPac(2) = DDraw.CreateSurfaceFromFile(fpImage + "pac\close_dn.bmp", SurfDesc1)
    Set sPac(3) = DDraw.CreateSurfaceFromFile(fpImage + "pac\close_lf.bmp", SurfDesc1)
    Set sPac(4) = DDraw.CreateSurfaceFromFile(fpImage + "pac\close_rg.bmp", SurfDesc1)
    Set sGhost(1) = DDraw.CreateSurfaceFromFile(fpImage + "ghost\gred_up.bmp", SurfDesc1)
    Set sGhost(2) = DDraw.CreateSurfaceFromFile(fpImage + "ghost\gcyan_up.bmp", SurfDesc1)
    Set sGhost(3) = DDraw.CreateSurfaceFromFile(fpImage + "ghost\ggreen_up.bmp", SurfDesc1)
    Set sGhost(4) = DDraw.CreateSurfaceFromFile(fpImage + "ghost\gyellow_up.bmp", SurfDesc1)
    
    Key.low = RGB(255, 255, 255)
    Key.high = RGB(255, 255, 255)

    sShield.SetColorKey DDCKEY_SRCBLT, Key
    Dim i As Integer
    For i = 1 To 4
        sPac(i).SetColorKey DDCKEY_SRCBLT, Key
        sGhost(i).SetColorKey DDCKEY_SRCBLT, Key
    Next i
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
    sShield.SetColorKey DDCKEY_SRCBLT, Key
End Sub

Sub Blt()
    '------Get Arena Location (RECT)------
    Dim aRect As RECT, aTop As Integer, aLeft As Integer
    DirectX.GetWindowRect frmLevelEditor.lvlSurface.hWnd, aRect
    aTop = aRect.Top
    aLeft = aRect.Left

    Dim rBack As RECT, rSprite As RECT, rTarget As RECT
    Dim X As Integer, Y As Integer, GhostNo As Integer, ChrDir As Integer
    Dim sTemp As DirectDrawSurface7
    
    rBack.Top = 0
    rBack.Left = 0
    rBack.Bottom = 418
    rBack.Right = 418
    sBack.Blt rBack, sBackground, rBack, DDBLT_WAIT
    
    rSprite.Bottom = 22: rSprite.Right = 22
        
    For Y = 0 To 18
        For X = 0 To 18
            Set sTemp = Nothing
            Select Case lvledBody.lvlSurf(X, Y)
                Case 1: Set sTemp = sFood
                Case 2: Set sTemp = sShield
                Case 3: Set sTemp = sWall
                Case 4: Set sTemp = sWall2
                Case Else: GoTo skipfor
            End Select
            
            rTarget.Top = Y * 22
            rTarget.Left = X * 22
            rTarget.Bottom = (Y + 1) * 22
            rTarget.Right = (X + 1) * 22
            sBack.Blt rTarget, sTemp, rSprite, DDBLT_KEYSRC Or DDBLT_WAIT
skipfor:
        Next X
    Next Y
    
    Dim PDir As Integer, Gno As Byte
    If frmLevelEditor.showChar.Checked Then
        With lvledBody.lvlPac
            If .xDir = 0 And .yDir = -1 Then PDir = 1
            If .xDir = 0 And .yDir = 1 Then PDir = 2
            If .xDir = -1 And .yDir = 0 Then PDir = 3
            If .xDir = 1 And .yDir = 0 Then PDir = 4
            rTarget.Top = .Y * 22
            rTarget.Left = .X * 22
            rTarget.Bottom = (.Y + 1) * 22
            rTarget.Right = (.X + 1) * 22
            sBack.Blt rTarget, sPac(PDir), rSprite, DDBLT_KEYSRC Or DDBLT_WAIT
        End With
        For Gno = 1 To 4
            With lvledBody.lvlGhost(Gno)
                rTarget.Top = .Y * 22
                rTarget.Left = .X * 22
                rTarget.Bottom = (.Y + 1) * 22
                rTarget.Right = (.X + 1) * 22
                sBack.Blt rTarget, sGhost(Gno), rSprite, DDBLT_KEYSRC Or DDBLT_WAIT
            End With
        Next Gno
    End If
    sPrimary.Blt aRect, sBack, rBack, DDBLT_WAIT
End Sub
