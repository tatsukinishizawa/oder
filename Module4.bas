Attribute VB_Name = "Module4"
Sub 社名()

Dim ws As Worksheet
Dim s As String
Dim i As Long, j As Long


Set ws = Worksheets("受注先")

s = "株式会社"

For j = 1 To 50

    Randomize
    ws.Cells(j, 1) = Int(50 * Rnd + 1) & s
    
Next j

    




End Sub
