Attribute VB_Name = "Module2"
Sub �N���A2()

Dim MB As Integer
Dim WS1 As Worksheet

Set WS1 = Worksheets("����")

MB = MsgBox("�����𑱍s���܂���", vbYesNo + vbExclamation, "���͓��e���N���A����܂�")


If MB = vbYes Then
WS1.Range("B7:G13").ClearContents

End If

If MB = vbNo Then Exit Sub
End Sub
