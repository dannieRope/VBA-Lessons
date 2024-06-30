Attribute VB_Name = "ModDoWhileDoUntil"
Option Explicit
Sub modDoUntil()
Range("A3").Select

Do Until ActiveCell.Value = ""
    ActiveCell.Offset(1, 0).Select
Loop
End Sub

Sub modDoWhile()
Range("A3").Select

Do While ActiveCell.Value <> ""
    ActiveCell.Offset(1, 0).Select
Loop
End Sub

Sub ModExistLoop()
Do
    If ActiveCell.Value = "" Then Exit Do
        ActiveCell.Offset(1, 0).Select
Loop
       
End Sub

Sub loopproject()
Dim FilmLength As Long
Dim filmRating As String

Range("A3").Select
Do Until ActiveCell.Value = ""

    FilmLength = ActiveCell.Offset(0, 3).Value
  
    If FilmLength < 100 Then
       filmRating = "Good"
    ElseIf FilmLength < 150 Then
       filmRating = "Very Good"
    Else: filmRating = "Excellent"
    End If
    ActiveCell.Offset(0, 6).Value = filmRating
    ActiveCell.Offset(1, 0).Select
       
 Loop
End Sub

Sub loopproject2()
Dim FilmLength As Long
Dim filmRating As String

Application.ScreenUpdating = False


Sheet1.Activate
Range("A3").Select
Do Until ActiveCell.Value = ""

    FilmLength = ActiveCell.Offset(0, 3).Value
  
    If FilmLength < 100 Then
       filmRating = "Good"
    ElseIf FilmLength < 150 Then
       filmRating = "Very Good"
    Else: filmRating = "Excellent"
    End If
    
    Range(ActiveCell, ActiveCell.End(xlToRight)).Copy
    Worksheets(filmRating).Activate
    ActiveCell.PasteSpecial
    ActiveCell.Offset(1, 0).Select
    
    Worksheets("VBA").Activate
    ActiveCell.Offset(1, 0).Select
       
 Loop
 Application.CutCopyMode = False
 Application.ScreenUpdating = True
End Sub
