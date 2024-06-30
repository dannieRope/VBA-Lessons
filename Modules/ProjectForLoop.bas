Attribute VB_Name = "ProjectForLoop"
Option Explicit
Sub filmRating()
Dim FilmLength As Long
Dim filmRating As String
Dim loopcount As Long
Dim cellCount As Long

cellCount = Range("A3", Range("A3").End(xlDown)).Count

For loopcount = 1 To cellCount
    FilmLength = Range("A2").Offset(loopcount, 3).Value
    
    If FilmLength < 100 Then
        filmRating = "Good"
    ElseIf FilmLength < 120 Then
        filmRating = "Very Good"
    Else: filmRating = "Excellent"
    End If

    Range("A2").Offset(loopcount, 6).Value = filmRating
    
Next loopcount

End Sub
