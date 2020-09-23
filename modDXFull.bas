Attribute VB_Name = "modDX"
Option Explicit
Public DirectX As New DirectX7
Public DDraw As DirectDraw7
Public sPrimary As DirectDrawSurface7
Public sBack As DirectDrawSurface7

Public sPac(1 To 2, 1 To 4) As DirectDrawSurface7
    'PacMan Character Surface (a,b):
    'a=1; Mouth Closed
    'a=2; Mouth Opened
    '-----------------
    'b=1; Face Up
    'b=2; Face Down
    'b=3; Face Left
    'b=4; Face Right
    
Public sGhost(1 To 5, 1 To 4) As DirectDrawSurface7
    'Ghost Character Surface (a,b):
    'a=1; Ghost Red
    'a=2; Ghost Blue/Cyan
    'a=3; Ghost Green
    'a=4; Ghost Yellow
    'a=5; Ghost sick
    '--------------------
    'b= Directional movement, same as pacman character
Public sBackground As DirectDrawSurface7
Public sFood As DirectDrawSurface7
Public sWall As DirectDrawSurface7
Public sWall2 As DirectDrawSurface7
    
Dim SurfDesc1 As DDSURFACEDESC2
Dim SurfDesc2 As DDSURFACEDESC2
Dim SurfDesc3 As DDSURFACEDESC2
Dim bRestore As Boolean
Public fpImage As String


Public aa, timee
Sub InitGame()
    fpImage = "C:\Windows\Desktop\img\"
    aa = 1
End Sub
Sub InitDirectX()
    Set DirectX = New DirectX7
    Set DDraw = DirectX.DirectDrawCreate("")
    frmMain.Show
    
    DDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    DDraw.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
         
    '------Init Primary Surface------
    With SurfDesc1
        .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
        .lBackBufferCount = 1
    End With
    Set sPrimary = DDraw.CreateSurface(SurfDesc1)
    
    SurfDesc1.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
    Set sBack = sPrimary.GetAttachedSurface(SurfDesc1.ddsCaps)
    sBack.SetForeColor QBColor(15)
    InitSurfaces
End Sub

Sub InitSurfaces()
    Set sPac(1, 1) = Nothing: Set sPac(1, 2) = Nothing: Set sPac(1, 3) = Nothing: Set sPac(1, 4) = Nothing
    Set sPac(2, 1) = Nothing: Set sPac(2, 2) = Nothing: Set sPac(2, 3) = Nothing: Set sPac(2, 4) = Nothing
    Set sGhost(1, 1) = Nothing: Set sGhost(1, 2) = Nothing: Set sGhost(1, 3) = Nothing: Set sGhost(1, 4) = Nothing
    Set sGhost(2, 1) = Nothing: Set sGhost(2, 2) = Nothing: Set sGhost(2, 3) = Nothing: Set sGhost(2, 4) = Nothing
    Set sGhost(3, 1) = Nothing: Set sGhost(3, 2) = Nothing: Set sGhost(3, 3) = Nothing: Set sGhost(3, 4) = Nothing
    Set sGhost(4, 1) = Nothing: Set sGhost(4, 2) = Nothing: Set sGhost(4, 3) = Nothing: Set sGhost(4, 4) = Nothing
    Set sGhost(5, 1) = Nothing
    
    '------Init Sprite Surface------
    '======Set Surface Description=====
    With SurfDesc2
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 22
        .lWidth = 22
    End With
    With SurfDesc3
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lHeight = 418
        .lWidth = 418
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
        
        Set sFood = .CreateSurfaceFromFile(fpImage + "schemes\5_food.bmp", SurfDesc2)
        Set sWall = .CreateSurfaceFromFile(fpImage + "schemes\5_wall.bmp", SurfDesc2)
        Set sWall2 = .CreateSurfaceFromFile(fpImage + "schemes\5_wall2.bmp", SurfDesc2)
        Set sBackground = DDraw.CreateSurfaceFromFile(fpImage + "schemes\5_Back.bmp", SurfDesc3)
    End With
    
    '=======Set Color Key for Sprites=======
    Dim Key As DDCOLORKEY
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
    sFood.SetColorKey DDCKEY_SRCBLT, Key
    sWall.SetColorKey DDCKEY_SRCBLT, Key
    sWall2.SetColorKey DDCKEY_SRCBLT, Key
End Sub

Sub Blt()
    Dim r1 As RECT
    Dim r2 As RECT
    Dim r3 As RECT
    Dim Y, X, delay

    timee = timee + 1
    If timee > 10 Then
        aa = aa * -1
        timee = 0
    End If
    

    r1.Top = 0
    r1.Left = 0
    r1.Right = 640
    r1.Bottom = 480

    sBack.BltColorFill r1, QBColor(0)
    
    r3.Bottom = 22: r3.Right = 22
            
    For X = 0 To 3
        If aa = -1 Then
            sBack.BltFast 0, X * 22, sPac(1, X + 1), r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Else
            sBack.BltFast 0, X * 22, sPac(2, X + 1), r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End If
    Next X
    
    For Y = 1 To 4
        For X = 0 To 3
            sBack.BltFast Y * 22, X * 22, sGhost(Y, X + 1), r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Next X
    Next Y
    sBack.BltFast 5 * 22, 0, sGhost(5, 1), r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    sBack.BltFast 0, 160, sFood, r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    sBack.BltFast 0, 190, sWall, r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    sBack.BltFast 0, 220, sWall2, r3, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    r2.Top = 0
    r2.Left = 0
    r2.Bottom = 418
    r2.Right = 418
    sBack.BltFast 200, 0, sBackground, r2, DDBLTFAST_WAIT
        
    sPrimary.Flip Nothing, DDFLIP_WAIT
End Sub

Function ExModeActive() As Boolean
    Dim TestCoopRes As Long
    TestCoopRes = DDraw.TestCooperativeLevel

    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
End Function

Sub EndIt()
DDraw.RestoreDisplayMode
DDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
End
End Sub
