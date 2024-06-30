Attribute VB_Name = "ModWithStatement"
Option Explicit
Sub Withstatement()
With Range("C3", Range("C3").End(xlDown))
    .Font.Italic = True
    .Font.Size = 11
    .Interior.Color = rgbAquamarine
    .NumberFormat = "dd/mm/yyyy"
End With

End Sub
