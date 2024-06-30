Attribute VB_Name = "ModLoop"
Option Explicit

Sub simpleloop()
Dim loopcounter As Integer

For loopcounter = 1 To 10
    Debug.Print loopcounter
Next loopcounter
End Sub

Sub simpleloopStep()
Dim loopcounter As Integer

For loopcounter = 1 To 10 Step 3
    Debug.Print loopcounter
Next loopcounter

End Sub

Sub reverseloopcounter()
Dim loopcounter As Integer

For loopcounter = 10 To 1 Step -1
    Debug.Print loopcounter
Next loopcounter
End Sub

Sub existloop()
Dim loopcounter As Integer
Dim RandomNum As Double

For loopcounter = 10 To 1 Step -1
    RandomNum = Math.Rnd
    If RandomNum < 0.1 Then Exit For
    Debug.Print loopcounter
Next loopcounter
End Sub

Sub loopProtectsheet()
Dim mysheet As Integer
Dim MaxSheetCount As Integer
MaxSheetCount = Worksheets.Count

For mysheet = 1 To MaxSheetCount
Worksheets(mysheet).Activate
Worksheets(mysheet).Protect
Next mysheet
End Sub

Sub loopUnProtectsheet()
Dim mysheet As Integer
Dim MaxSheetCount As Integer
MaxSheetCount = Worksheets.Count

For mysheet = 1 To MaxSheetCount
Worksheets(mysheet).Activate
Worksheets(mysheet).Unprotect
Next mysheet
End Sub

Sub loopworksheets()
'loop to close workbooks
Dim myworkbook As Integer
Dim WorkbookCount As Integer
WorkbookCount = Workbooks.Count

For myworkbook = 2 To WorkbookCount
'alt For myworkbook = WorkbookCount To 2 Step -1
    Workbooks(myworkbook).Close
Next myworkbook
End Sub


