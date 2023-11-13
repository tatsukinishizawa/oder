Attribute VB_Name = "Module4"
Sub Ğ–¼()

Dim ws As Worksheet
Dim s As String
Dim i As Long, j As Long


Set ws = Worksheets("ó’æ")

s = "Š”®‰ïĞ"

For j = 1 To 50

    Randomize
    ws.Cells(j, 1) = Int(50 * Rnd + 1) & s
    
Next j

    




End Sub
