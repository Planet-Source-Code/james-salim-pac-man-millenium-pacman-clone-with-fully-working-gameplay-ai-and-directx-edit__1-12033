Attribute VB_Name = "modGameHighScore"
Type HighScoreList
    Name As String * 10
    LastLevel As Integer
    Score As Long
End Type

Public HScoreList(10) As HighScoreList
Public HScoreMinValue As Long
Public HScoreMaxValue As Long
Dim Filenumber As Long

Sub LoadHScore()
    Filenumber = FreeFile

    Open fpGame + "High.dat" For Random As #Filenumber Len = 16
        For i = 1 To 10
            Get #Filenumber, i, HScoreList(i)
        Next i
    Close #1
    GetMinimumScoreNo
End Sub

Sub GetMinimumScoreNo()
    HScoreMinValue = HScoreList(1).Score
    HScoreMaxValue = 0
    For i = 1 To 10
        If HScoreMinValue > HScoreList(i).Score Then HScoreMinValue = HScoreList(i).Score
        If HScoreMaxValue < HScoreList(i).Score Then HScoreMaxValue = HScoreList(i).Score
    Next i
End Sub

Sub SaveHScore(Name As String, LastLevel As Integer, Score As Long)
    Filenumber = FreeFile
    Dim TempHScore As HighScoreList
    
    For i = 1 To 10
        If Score > HScoreList(i).Score Then ScoreRank = i: Exit For
    Next i
    
    Open fpGame + "High.dat" For Random As #Filenumber Len = 16
        For i = 10 To ScoreRank Step -1
            Get #Filenumber, i, TempHScore
            Put #Filenumber, i + 1, TempHScore
        Next i
        TempHScore.Name = Name
        TempHScore.LastLevel = LastLevel
        TempHScore.Score = Score
        Put #Filenumber, ScoreRank, TempHScore
    Close #1
End Sub
