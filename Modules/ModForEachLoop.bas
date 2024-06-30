Attribute VB_Name = "ModForEachLoop"
Option Explicit

Sub ListWorksheetNames()
Dim sheetnumber As Integer
Dim sheetName As String
Dim list As Range

Set list = Range("A1")
For sheetnumber = 1 To Worksheets.Count
    sheetName = Worksheets(sheetnumber).Name
    list.Offset(sheetnumber).Value = sheetName
    Debug.Print sheetName
Next sheetnumber

End Sub

Sub ListWorksheetNames2()
Dim singlesheets As Worksheet

For Each singlesheets In Worksheets
Debug.Print singlesheets.Name
Next singlesheets

End Sub
Sub forloopeachprotect()
Dim mysheet As Worksheet

For Each mysheet In Worksheets
    mysheet.Protect ("123")
Next mysheet
End Sub

Sub forloopeachUnprotect()
Dim mysheet As Worksheet

For Each mysheet In Worksheets
    mysheet.Unprotect
Next mysheet
End Sub

Sub listworkbooks()
Dim myworkbook As Workbook
Dim BookName As String

For Each myworkbook In Workbooks
  BookName = myworkbook.Name
  Debug.Print BookName
Next myworkbook
End Sub

Sub closeallworkbooks()

Dim myworkbook As Workbook
Dim BookName As String

For Each myworkbook In Workbooks
  BookName = myworkbook.Name
  'alt = if thisWorkBook.name <> "VBA Lessons.xlsm" then ..
  If BookName <> "VBA Lessons.xlsm" Then
    myworkbook.Close
  End If
Next myworkbook
End Sub

Sub foreachcharts()
Dim mychart As ChartObject

Set mychart = Worksheets("VBA").ChartObjects(3)
mychart.Chart.SetSourceData Range("B2:B20,D2:D20")

End Sub

Sub foreachcharts1()
    Dim mychart As ChartObject
    For Each mychart In Sheet1.ChartObjects
        mychart.Chart.SetSourceData Range("B2:B20,D2:D20")
    Next mychart
End Sub
Sub listcellvaluesintitle()
Dim titles As Range
Dim cell As Range

 Set titles = Range("A3", Range("A3").End(xlDown))

For Each cell In titles
    If cell.Offset(0, 3).Value > 120 Then
        Debug.Print cell.Offset(0, 1).Value
    End If
Next cell

End Sub

Sub listfilmtitleinnewworkbook()
Dim titles As Range
Dim cell As Range

ThisWorkbook.Activate
Sheet1.Activate

Set titles = Range("A3", Range("A3").End(xlDown))
 
Workbooks.Add
Range("A1").Value = "List of Film Titles"

Range("A2").Select

For Each cell In titles
    If cell.Offset(0, 3).Value > 120 Then
        ActiveCell.Value = cell.Offset(0, 1).Value
        ActiveCell.Offset(1).Select
    End If
Next cell

End Sub


