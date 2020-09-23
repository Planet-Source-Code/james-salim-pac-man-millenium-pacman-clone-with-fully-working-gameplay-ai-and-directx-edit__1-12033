Attribute VB_Name = "modLevelEditor"
Type LevelPacChar
    StartDelay As Integer:    DrunkDelay As Integer
    X As Byte
    Y As Byte
    xDir As Integer
    yDir As Integer
    ProtectTime As Integer
    DrunkTime As Integer
End Type

Type LevelGhostChar
    X As Byte
    Y As Byte
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
 
Public Type LevelScheme
    Back As Byte
    Food As Byte
    Wall1 As Byte
    Wall2 As Byte
End Type

Type LevelBody
    LvlNo As Integer
    lvlName As String * 32
    lvlSurf(18, 18) As Byte
    lvlScheme As LevelScheme
    lvlPac As LevelPacChar
    lvlGhost(1 To 4) As LevelGhostChar
    lvlItems As LevelItems
End Type

Public lvledBody As LevelBody
Public curItem As Integer
Public fpImage As String
Public fpGame As String
Public MaxLevelNo As Integer

Public lvledpropItemSelected As Integer
Public lvledpropGhostSelected As Integer
    
Sub Init()
    fpImage = App.Path + "\pacman\image\"
    fpGame = App.Path + "\pacman\game\"
End Sub

Sub lvledRefreshSchemes()
    With frmLevelEditor
        .Toolbar.ImageList = Nothing

        .SchemeList.ListImages.Add 1, , LoadPicture(fpImage + "schemes\" + LTrim$(Str$(lvledBody.lvlScheme.Food)) + "_food.bmp")
        .SchemeList.ListImages.Add 2, , LoadPicture(fpImage + "item\protect.bmp")
        .SchemeList.ListImages.Add 3, , LoadPicture(fpImage + "schemes\" + LTrim$(Str$(lvledBody.lvlScheme.Wall1)) + "_wall.bmp")
        .SchemeList.ListImages.Add 4, , LoadPicture(fpImage + "schemes\" + LTrim$(Str$(lvledBody.lvlScheme.Wall2)) + "_wall2.bmp")
        .SchemeList.ListImages.Add 5, , LoadPicture(fpImage + "schemes\erase.ico")
        
        .Toolbar.ImageList = .SchemeList.Object
        .Toolbar.buttons(1).Image = 1
        .Toolbar.buttons(2).Image = 2
        .Toolbar.buttons(3).Image = 3
        .Toolbar.buttons(4).Image = 4
        .Toolbar.buttons(5).Image = 5
    End With
    InitLevelSurfaces lvledBody.lvlScheme
End Sub

Sub lvledSetSurface(X, Y, SurfNo)
    lvledBody.lvlSurf(X, Y) = SurfNo
End Sub

Sub lvledInitLevelView()
    Dim Filenumber As Long
    Dim Tempbody As LevelBody
    Filenumber = FreeFile
    
    On Error GoTo ErrorProc

    Open fpGame + "level.dat" For Random Access Read As #Filenumber Len = 461
        Get #Filenumber, 1, Tempbody
        MaxLevelNo = Tempbody.LvlNo
                
        With frmLevelEditor
            .lvlView.Nodes.Clear

            For lvl = 1 To MaxLevelNo
                Get #Filenumber, lvl + 1, Tempbody
                .lvlView.Nodes.Add , , , Tempbody.lvlName
            Next lvl
            .mnuLevel.Enabled = False
        End With
    Close #Filenumber
    Exit Sub
    
ErrorProc:
End Sub

Sub lvledLoadLevel(LvlNo As Integer)
    Filenumber = FreeFile
    Dim Tempbody As LevelBody
    On Error GoTo ErrorProc
    
    If LvlNo = 0 Then
         lvledBody = CreateEmptyLevel
    Else
        Open fpGame + "level.dat" For Random Access Read As #Filenumber Len = 461
            If LvlNo > MaxLevelNo Or LvlNo < 0 Then GoTo ErrorProc
            Get #Filenumber, LvlNo + 1, Tempbody
            lvledBody = Tempbody
        Close #Filenumber
    End If
    
    lvledRefreshSchemes
ErrorProc:
End Sub

Sub lvledSaveLevel()
    Filenumber = FreeFile
    On Error GoTo ErrorProc
    
    Open fpGame + "level.dat" For Random As #Filenumber Len = 461
        Put #Filenumber, SelectedLevel + 1, lvledBody
    Close #Filenumber
    lvledInitLevelView
ErrorProc:
End Sub

Function SelectedLevel() As Integer
    With frmLevelEditor.lvlView
        For LvlNo = 1 To .Nodes.Count
            If .Nodes(LvlNo).Selected Then SelectedLevel = LvlNo: Exit Function
        Next LvlNo
    End With
End Function

Sub InsertLevel(Above As Boolean)
    Dim Tempbody As LevelBody
    On Error GoTo ErrorProc
    
    Filenumber = FreeFile
    Open fpGame + "level.dat" For Random As #Filenumber Len = 461
        Get #Filenumber, 1, Tempbody
        Tempbody.LvlNo = Tempbody.LvlNo + 1
        Put #Filenumber, 1, Tempbody
        
        If Not (Above) Then AdditionalLevel = 1 Else AdditionalLevel = 0
        For lvl = Tempbody.LvlNo + 1 To SelectedLevel + AdditionalLevel Step -1
            Get #Filenumber, lvl + 1, Tempbody
            Put #Filenumber, lvl + 2, Tempbody
        Next lvl
        Put #Filenumber, SelectedLevel + AdditionalLevel + 1, CreateEmptyLevel
    Close #Filenumber
    lvledInitLevelView
ErrorProc:
End Sub

Sub DeleteLevel()
    Filenumber = FreeFile
    Dim Tempbody As LevelBody
    On Error GoTo ErrorProc
    
    Open fpGame + "level.dat" For Random As #Filenumber Len = 461
        Get #Filenumber, 1, Tempbody
        Tempbody.LvlNo = Tempbody.LvlNo - 1
        Put #Filenumber, 1, Tempbody
        
        For lvl = SelectedLevel + 1 To Tempbody.LvlNo + 1
            Get #Filenumber, lvl + 1, Tempbody
            Put #Filenumber, lvl, Tempbody
        Next lvl
    Close #Filenumber
    lvledInitLevelView
ErrorProc:
End Sub

Function CreateEmptyLevel() As LevelBody
Dim Tempbody As LevelBody
With Tempbody
    .lvlName = "<untitled>"
    .LvlNo = 0
            
    '--Pac--
    With .lvlPac
        .DrunkDelay = 0: .StartDelay = 0
        .DrunkTime = 0: .ProtectTime = 0
        .X = 0: .Y = 0: .xDir = 1: .yDir = 0
    End With
    
    '--Ghost--
    For Gno = 1 To 4
        With .lvlGhost(Gno)
            .Delay = 0: .SickDelay = 0
            .X = 0: .Y = 0
        End With
    Next Gno
            
    '--Item--
    Dim TempItem As LevelItemChar
    With TempItem
        .Amount = 0
        .AppearTime = 0
        .Delay = 0
    End With
    .lvlItems.Beer = TempItem
    .lvlItems.Berry = TempItem
    .lvlItems.Cherry = TempItem
    .lvlItems.Life = TempItem
    
    '--Scheme--
    With .lvlScheme
        .Back = 0
        .Food = 0
        .Wall1 = 0
        .Wall2 = 0
    End With
    
    '--Surface--
    For Y = 0 To 18
        For X = 0 To 18
            .lvlSurf(X, Y) = 0
        Next X
    Next Y
End With
CreateEmptyLevel = Tempbody
End Function

Sub SpreadSurface(ItemNo As Integer)
    For Y = 0 To 18
        For X = 0 To 18
                lvledSetSurface X, Y, ItemNo
        Next X
    Next Y
End Sub
