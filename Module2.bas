Attribute VB_Name = "Module2"
Sub クリア2()

Dim MB As Integer
Dim WS1 As Worksheet

Set WS1 = Worksheets("入力")

MB = MsgBox("処理を続行しますか", vbYesNo + vbExclamation, "入力内容がクリアされます")


If MB = vbYes Then
WS1.Range("B7:G13").ClearContents

End If

If MB = vbNo Then Exit Sub
End Sub
