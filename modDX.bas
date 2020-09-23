Attribute VB_Name = "modDX"
'=====================================================================
'=DIRECTX DECLARATION=================================================
'=====================================================================
Public DirectX As New DirectX7
Public DDraw As DirectDraw7

Public sTitleScreen As DirectDrawSurface7
Public sPause As DirectDrawSurface7
Public sTemp As DirectDrawSurface7

Public sPrimary As DirectDrawSurface7
Public sBack  As DirectDrawSurface7
Public ddClipper As DirectDrawClipper

Public sPac(1 To 2, 1 To 4) As DirectDrawSurface7
Public sPacDead As DirectDrawSurface7
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
    
Public sItemShield As DirectDrawSurface7
Public sItemLife As DirectDrawSurface7
Public sItemBerry As DirectDrawSurface7
Public sItemCherry As DirectDrawSurface7
Public sItemBeer As DirectDrawSurface7
Public sPacShield As DirectDrawSurface7
    
Public sBackground As DirectDrawSurface7    'Level schemes background surface
Public sFood As DirectDrawSurface7          'Level schemes food surface
Public sWall As DirectDrawSurface7          'Level schemes wall surface
Public sWall2 As DirectDrawSurface7         'Level schemes 2nd wall surface
    
Public aRect As RECT
Public aTop As Long
Public aLeft As Long

Sub InitDirectX()
InitDirectDraw
End Sub
