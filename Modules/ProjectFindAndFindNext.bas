Attribute VB_Name = "ProjectFindAndFindNext"
Option Explicit

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
