Attribute VB_Name = "ModFindandFindNext"
Option Explicit

Sub FindFillm()

Range("B3:B22").Find("The").Select

End Sub

Sub FindFilmWholesearch()
Range("B3:B22").Find(What:="The Lorax", LookAt:=xlWhole).Select
End Sub

Sub findCaseSensitive()
Range("B3:B22").Find(What:="The Lorax", matchcase:=True).Select
End Sub

Sub FilmNotFound()
Dim FilmSearch As Range
Set FilmSearch = Range("B3:B22").Find(What:="The Skyfall", matchcase:=True)
If FilmSearch Is Nothing Then
    MsgBox "Film not Found"
Else: FilmSearch.Select
End If
End Sub

Sub findfilmProject()
Dim FilmName As String
Dim FilmRange As Range
Dim FilmSearch As Range
Dim FirstFilmCell As String

FilmName = InputBox("Enter Film Name")
Set FilmRange = Range("B3", Range("B3").End(xlDown))

Set FilmSearch = FilmRange.Find(What:=FilmName, LookAt:=xlPart, matchcase:=False)

If FilmSearch Is Nothing Then
    MsgBox FilmName & " Not Found"
Else
    FirstFilmCell = FilmSearch.Address
    Do
    MsgBox FilmSearch.Value & " released on " & FilmSearch.Offset(0, 1).Value
    Set FilmSearch = FilmRange.FindNext(FilmSearch)
    Loop While FilmSearch.Address <> FirstFilmCell
End If
End Sub

