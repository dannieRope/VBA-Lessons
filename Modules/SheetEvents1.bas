
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim myrange As Range
    Set myrange = Intersect(Target, Range("A1:C6"))
    
    If Not myrange Is Nothing Then
      myrange.Interior.Color = rgbAliceBlue
    End If
'*****************************************************
'Alternate way
'*****************************************************
    If Target.Row = 6 And Target.Column < 4 Then
       Target.Interior.Color = rgbAliceBlue
    End If
    

End Sub
