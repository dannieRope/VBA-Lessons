Attribute VB_Name = "ModIfStatement"
Option Explicit
Sub ifstatement()
Dim MovieLen As Range
Dim cell As Range

Set MovieLen = Range("D3", Range("D3").End(xlDown))
For Each cell In MovieLen
    If cell.Value < 100 Then
        cell.Offset(0, 2).Value = "Short"
    Else: cell.Offset(0, 2).Value = "Long"
    End If
Next cell

End Sub


