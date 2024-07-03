Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Hour(Now) < 8 Then
    MsgBox "You're not leaving"
    Cancel = True
    End If
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    Dim howmany As Integer
    
    If TypeOf Sh Is Worksheet Then
    howmany = InputBox("Number of sheets to add")
    
    Application.EnableEvents = False
    Worksheets.Add Count:=howmany - 1
    Application.EnableEvents = True
    
    End If
    

End Sub

Private Sub Workbook_Open()

MsgBox "Welcome!, Your VBA Journey Begins here", , Date

End Sub
