Attribute VB_Name = "Module4"
Sub �Ж�()

Dim ws As Worksheet
Dim s As String
Dim i As Long, j As Long


Set ws = Worksheets("�󒍐�")

s = "�������"

For j = 1 To 50

    Randomize
    ws.Cells(j, 1) = Int(50 * Rnd + 1) & s
    
Next j

    




End Sub
