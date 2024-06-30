Attribute VB_Name = "ProjectFindAMovie"
Option Explicit

Sub findarange()
Dim FilmtoFind As String
Dim FilmCell As Range

FilmtoFind = InputBox("Type in a Film")
Set FilmCell = _
    Range("B3", Range("B3").End(xlDown)).Find(FilmtoFind)

If FilmCell Is Nothing Then
    MsgBox FilmtoFind & " Not found"
Else: MsgBox FilmCell.Value & " Was found in " & FilmCell.Address
End If

End Sub

