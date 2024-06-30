Attribute VB_Name = "ModInputBox"
Option Explicit
Sub whatisyourname()
Dim myname As String
myname = InputBox("Please type in your name", "Personal Details")

If myname = "" Then
    MsgBox "You didn't enter your name", vbExclamation
Else: MsgBox "Hi Welcome " & myname
End If

End Sub
