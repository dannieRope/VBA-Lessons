Attribute VB_Name = "ModDeclareVar"
Option Explicit
Dim Title As String
Dim ReleaseDate As Date

Sub GetUserInput()
Title = InputBox("Enter Move Title")
ReleaseDate = InputBox("Enter Release Date")
Sheets("VBA").Range("A1").End(xlDown).Offset(1).Select

Call AddFilmToList

End Sub

Sub AddFilmToList()
ActiveCell.Value = ActiveCell.Offset(-1).Value + 1
ActiveCell.Offset(0, 1).Value = Title
ActiveCell.Offset(0, 2).Value = ReleaseDate

MsgBox Title & " Was added to the list"
End Sub

