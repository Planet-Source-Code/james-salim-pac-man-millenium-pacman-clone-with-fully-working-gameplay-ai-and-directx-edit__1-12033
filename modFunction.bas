Attribute VB_Name = "modFunction"
Function CheckWall(x, y) As Boolean
    CheckWall = False
    If x < 0 Then x = MaxGameX
    If y < 0 Then y = MaxGameY
    If x > MaxGameX Then x = 0
    If y > MaxGameY Then y = 0
    If aGame(x, y) = pac_Wall Or aGame(x, y) = pac_Wall2 Then CheckWall = True
    'If aGame2(x, y) >= pac_Ghost(1, 1) And aGame2(x, y) <= pac_Ghost(5, 4) Then CheckWall = True
End Function

Sub GhostAddMove(GNo, x, y)
DoEvents
For i = 0 To 9
    Ghost(GNo).LastX(i + 1) = Ghost(GNo).LastX(i)
    Ghost(GNo).LastY(i + 1) = Ghost(GNo).LastY(i)
Next i
Ghost(GNo).LastX(0) = x
Ghost(GNo).LastY(0) = y
End Sub

Function GhostPast(GNo, x, y) As Integer
    DoEvents
    Past = 0
    With Ghost(GNo)
        For i = 0 To 10
            If .LastX(i) = x And .LastY(i) = y Then Past = Past + 1
        Next i
        GhostPast = Past
    End With
End Function

Function GetMaxValue(Val1, Val2, Val3, Val4) As Integer
    If Val1 >= Val2 Then MaxValue = Val1 Else MaxValue = Val2
    If Val3 > MaxValue Then MaxValue = Val3
    If Val4 > MaxValue Then MaxValue = Val4
    GetMaxValue = MaxValue
End Function

Function GetAmountEqualTo(Val1, Val2, Val3, Val4, CompareVal) As Integer
    If Val1 = CompareVal Then AmountEqual = AmountEqual + 1
    If Val2 = CompareVal Then AmountEqual = AmountEqual + 1
    If Val3 = CompareVal Then AmountEqual = AmountEqual + 1
    If Val4 = CompareVal Then AmountEqual = AmountEqual + 1
    GetAmountEqualTo = AmountEqual
End Function

Function PixelsToTwips_height(pxls)
    PixelsToTwips_height = pxls * (frmMain.Height / frmMain.ScaleHeight)
End Function

Function PixelsToTwips_width(pxls)
    PixelsToTwips_width = pxls * (frmMain.Width / frmMain.ScaleWidth)
End Function

Sub SetViewMode(ViewNo As Integer)
    Dim ArenaX As Integer, ArenaY As Integer
    Select Case ViewNo
        Case 0
            ArenaX = 418: ArenaY = 418
        Case 1
            ArenaX = 627: ArenaY = 627
        Case 2
            ArenaX = 836: ArenaY = 836
    End Select
    With frmMain
        For i = 1 To 4
            .Height = PixelsToTwips_height(ArenaY + 1) + 225
            .Width = PixelsToTwips_width(ArenaX + 1)
        Next i
        .Arena.Width = ArenaX
        .Arena.Height = ArenaY
    End With
End Sub
