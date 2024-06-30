Attribute VB_Name = "ModMsgBox"
Option Explicit

Sub simpleMessage()
Dim ButtonClicked As VbMsgBoxResult
MsgBox "Hi World", vbInformation, "Welcome Message"
'Same As
MsgBox prompt:="Hi World", Buttons:=vbInformation, Title:="Welcome Message"

ButtonClicked = MsgBox("Do you line Pizza", vbQuestion + vbYesNo, "Food Question")

If ButtonClicked = vbYes Then
  MsgBox "Yes, Pizza's are great!", vbExclamation
Else: MsgBox "Why not? Pizza's are great!", vbCritical
End If


End Sub

Sub DateMessage()
MsgBox "The Date is " & Date & vbNewLine & "Weather Looks Sunny"
End Sub

Sub MovieMessage()
Range("B9").Select
MsgBox ActiveCell.Value & " was released on " & ActiveCell.Offset(0, 1).Value
'same as
MsgBox Selection.Value & " was released on " & Selection.Offset(0, 1).Value
End Sub


