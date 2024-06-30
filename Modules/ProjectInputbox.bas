Attribute VB_Name = "ProjectInputbox"
Option Explicit
Sub Inputfilmdetail()
Dim Title As String

Dim strMovieDate As String
Dim datMovieDate As Date

Dim Length As Integer
Dim strLength As String

Title = InputBox("Enter Movie Title")

strMovieDate = InputBox("Enter Movie Release Date")

If strMovieDate = "" Then
    MsgBox "You did enter Valid Date", vbExclamation
    Exit Sub
End If

strLength = InputBox("Enter Movie Length in Mins")
If strLength = "" Then
    MsgBox "Enter valid movie Length"
    Exit Sub
End If

datMovieDate = CDate(strMovieDate)
Length = CInt(strLength)

Range("A3").End(xlDown).Offset(1).Select
Selection.Value = Selection.Offset(-1).Value + 1
Selection.Offset(0, 1).Value = Title
Selection.Offset(0, 2).Value = datMovieDate
Selection.Offset(0, 3).Value = Length

End Sub

Sub applicationinputbox()
Dim Title As String
Dim MovieDate As Date
Dim Length As Integer

Title = Application.InputBox("Enter Movie Title")
MovieDate = Application.InputBox(prompt:="Enter Release Date (dd/mm/yyyy)", Type:=1)
Length = Application.InputBox(prompt:="Enter Movie Length", Type:=1)


Range("A3").End(xlDown).Offset(1).Select
Selection.Value = Selection.Offset(-1).Value + 1
Selection.Offset(0, 1).Value = Title
Selection.Offset(0, 2).Value = MovieDate
Selection.Offset(0, 3).Value = Length

End Sub

Sub EnterAformula()
Dim myformula As String
myformula = Application.InputBox(prompt:="Enter a Formula", Type:=0)
Range("G2").FormulaLocal = myformula
End Sub
Sub enterFormulaandRange()
Dim myformula As String
Dim Formulacell As Range

myformula = Application.InputBox(prompt:="Enter Formula", Type:=0)

Set Formulacell = Application.InputBox(prompt:="Select Result cell", Type:=8)

Formulacell.FormulaLocal = myformula

End Sub

Sub copypaste()
Dim CopyRange As Range
Dim PasteRange As Range

Set CopyRange = Application.InputBox(prompt:="Select Range to Copy", Type:=8)
Set PasteRange = Application.InputBox(prompt:="Select Range to paste", Type:=8)

CopyRange.Copy PasteRange
End Sub

Sub ReturnArray()
Dim FilmLength() As Variant
Dim loopcounter As Long
Dim ResultRange As Range

FilmLength = Application.InputBox(prompt:="Choose length to convert", Type:=64)

For loopcounter = LBound(FilmLength, 1) To UBound(FilmLength, 1)

    FilmLength(loopcounter, 1) = Int(FilmLength(loopcounter, 1) / 60) & " Hour " _
        & FilmLength(loopcounter, 1) Mod 60 & " Mins"
        
Next loopcounter

Set ResultRange = _
    Application.InputBox(prompt:="Choose Result cells", Type:=8)
    
Set ResultRange = _
    Range(ResultRange, ResultRange.Offset(UBound(FilmLength, 1) - 1, 0))
    
ResultRange = FilmLength

    
End Sub
