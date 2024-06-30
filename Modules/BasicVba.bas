Attribute VB_Name = "BasicVba"
Option Explicit

Sub createAndLabelNewSheet()
Range("A1").Value = "Created by"
Range("A1").Offset(0, 1).Value = Environ("UserName")
Range("A2").Value = "Created On"
Range("A2").Offset(0, 1).Value = Now()


With Range("A1:A2")
       .Font.Color = vbBlue
       .Interior.Color = rgbLightCyan
End With

End Sub

Sub selectionofcells()
Cells(13, 2).Select
ActiveCell.Value = 10 'hold ctrl + space to activate intellices
End Sub


Sub nameranges()
Range("Title").Font.Color = vbBlue
[Title].Font.Italic = True
[Release_Date].Font.Italic = True

End Sub

Sub lastrow1()

Range("ID").End(xlDown).Offset(1).Value = WorksheetFunction.Max(Range("ID").CurrentRegion.Columns(1)) + 1
Range("ID").End(xlDown).Offset(1).Value = 1 + WorksheetFunction.Max(Range("A1").CurrentRegion.Columns(1))

End Sub

Sub lastrow2()
Range("ID").End(xlDown).Offset(1).Select
ActiveCell.Value = ActiveCell.Offset(-1).Value + 1
ActiveCell.Offset(0, 1).Value = "300:The rise of an Empire"
ActiveCell.Offset(0, 2).Value = #5/3/2021#
End Sub

Sub selectdata()
Range("A3", Range("A1").End(xlDown)).Select
Selection.Font.Italic = True
Range("B3", Range("B3").End(xlDown)).Select
Range("B3", Range("B3").End(xlDown)).Font.Color = rgbBlue

'selecting and entire block of data
Range("A3", Range("A1").End(xlDown).End(xlToRight)).Select
Selection.Interior.Color = rgbAliceBlue

End Sub

Sub selectusingcurrentregion()
Sheet4.Activate
Range("A1").CurrentRegion.Select
Selection.Copy
Sheets("VBA").Activate
Range("A1").PasteSpecial xlPasteColumnWidths

Application.CutCopyMode = False


End Sub

Sub selectusingcurrentregion2()
Range("A1").CurrentRegion.Select
Selection.Copy Range("A20")
Application.CutCopyMode = False

End Sub
Sub copyandpastedata()
Sheets("VBA").Range("A1").CurrentRegion.Copy
Sheet4.Activate
Sheet4.Range("A1").PasteSpecial
Columns("A:C").AutoFit

End Sub
