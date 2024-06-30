Attribute VB_Name = "ModObjectVar"
Option Explicit

Sub storeRangeCells()
Dim FilmTitleCells As Range
Set FilmTitleCells = Range("B3", Range("B3").End(xlDown))
FilmTitleCells.Font.Color = rgbBlue
FilmTitleCells.Font.Italic = True

End Sub

Sub referencingWorksheetObjects()
Dim MyNewSheet As Worksheet

Worksheets.Add

Set MyNewSheet = ActiveSheet
Sheets("VBA").Range("A1").CurrentRegion.Copy
MyNewSheet.Activate
ActiveCell.PasteSpecial
End Sub
Sub alternatetive()
Dim MyNewSheet As Worksheet

Set MyNewSheet = Worksheets.Add

Sheets("VBA").Range("A1").CurrentRegion.Copy
MyNewSheet.Activate
ActiveCell.PasteSpecial
End Sub

Sub others()
Dim MyNewBook As Workbook
Set MyNewBook = Workbooks.Add

Dim MyNewChart As Chart
Set MyNewChart = Charts.Add

End Sub

Sub findarange()
Dim FilmtoFind As String
Dim FilmCell As Range

FilmtoFind = InputBox("Type in a Film")
Set FilmCell = _
    Range("B3", Range("B3").End(xlDown)).Find(FilmtoFind)

If FilmCell Is Nothing Then
    MsgBox FilmCell.Value & " Not found"
Else: MsgBox FilmCell.Value & " Was found in " & FilmCell.Address
End If

End Sub
