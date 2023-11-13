Attribute VB_Name = "Module1"

Sub 注文書()


Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim LCRN1 As Long
Dim AllRow As Long

AllRow = ActiveSheet.Rows.Count '全ての行数

Set WS1 = Worksheets("入力")
Set WS2 = Worksheets("検索リスト")

'未入力回避

If WS1.Range("E4") = "" Then

    MsgBox "識別が入力されていません"
    Exit Sub

End If


Dim WS4 As Worksheet
Dim LCRN4 As Long
Dim LCRN6 As Long

Set WS4 = Worksheets("発注番号")
LCRN4 = WS4.Cells(AllRow, 1).End(xlUp).Row + 1
LCRN6 = WS4.Cells(Rows.Count, "A").End(xlUp)

Dim RefNum As String '発注番号振り分け

 RefNum = LCRN6 + 1

WS4.Cells(LCRN4, 1) = RefNum
WS4.Cells(LCRN4, 2) = Date
WS4.Cells(LCRN4, 3) = WS1.Range("C3")
WS4.Cells(LCRN4, 4) = WS1.Range("E4")
WS4.Cells(LCRN4, 5) = WS1.Range("B7")
WS4.Cells(LCRN4, 6) = WS1.Range("C7")
WS4.Cells(LCRN4, 7) = WS1.Range("F7")


Dim WB1 As Workbook
Set WB1 = ActiveWorkbook



'見積書の原紙をコピーする

Worksheets("注文書").Copy
ActiveSheet.Name = "注文書"

Dim WS5 As Worksheet
Dim WB2 As Workbook

Set WB2 = ActiveWorkbook
Set WS5 = ActiveSheet


Dim WS6 As Worksheet

Set WS6 = Sheets.Add(After:=WS5)

WS6.Name = ("受注先")

WS6.Range("A1") = WS2.Range("A1")
WS6.Range("A2") = WS2.Range("A2")
WS6.Range("B1") = WS2.Range("B1")
WS6.Range("B2") = WS2.Range("B2")



'入力シートのデータを、注文書に転記していく


WS5.Range("A15").MergeArea = RefNum
WS5.Range("A5").MergeArea.Formula = "=(受注先!A1)"


WS5.Range("B7").MergeArea = WS1.Range("E3") '項目
WS5.Range("J3").MergeArea = Date

WS5.Range("B15") = WS1.Range("E4")

WS5.Range("D15").MergeArea = WS1.Range("B7") '品名
WS5.Range("D16").MergeArea = WS1.Range("B8")
WS5.Range("D17").MergeArea = WS1.Range("B9")
WS5.Range("D18").MergeArea = WS1.Range("B10")
WS5.Range("D19").MergeArea = WS1.Range("B11")
WS5.Range("D20").MergeArea = WS1.Range("B12")
WS5.Range("D21").MergeArea = WS1.Range("B13")
WS5.Range("D22").MergeArea = WS1.Range("B14")
WS5.Range("D23").MergeArea = WS1.Range("B15")
WS5.Range("D24").MergeArea = WS1.Range("B16")
WS5.Range("D25").MergeArea = WS1.Range("B17")
WS5.Range("D26").MergeArea = WS1.Range("B18")

WS5.Range("F15") = WS1.Range("C7") '型式
WS5.Range("F16") = WS1.Range("C8")
WS5.Range("F17") = WS1.Range("C9")
WS5.Range("F18") = WS1.Range("C10")
WS5.Range("F19") = WS1.Range("C11")
WS5.Range("F20") = WS1.Range("C12")
WS5.Range("F21") = WS1.Range("C13")
WS5.Range("F22") = WS1.Range("C14")
WS5.Range("F23") = WS1.Range("C15")
WS5.Range("F24") = WS1.Range("C16")
WS5.Range("F25") = WS1.Range("C17")
WS5.Range("F26") = WS1.Range("C18")
WS5.Range("F27") = WS1.Range("C19")

WS5.Range("G15") = WS1.Range("D7") '材質
WS5.Range("G16") = WS1.Range("D8")
WS5.Range("G17") = WS1.Range("D9")
WS5.Range("G18") = WS1.Range("D10")
WS5.Range("G19") = WS1.Range("D11")
WS5.Range("G20") = WS1.Range("D12")
WS5.Range("G21") = WS1.Range("D13")
WS5.Range("G22") = WS1.Range("D14")
WS5.Range("G23") = WS1.Range("D15")
WS5.Range("G24") = WS1.Range("D16")
WS5.Range("G25") = WS1.Range("D17")
WS5.Range("G26") = WS1.Range("D18")
WS5.Range("G27") = WS1.Range("D19")

WS5.Range("H15") = WS1.Range("E7") '数量
WS5.Range("H16") = WS1.Range("E8")
WS5.Range("H17") = WS1.Range("E9")
WS5.Range("H18") = WS1.Range("E10")
WS5.Range("H19") = WS1.Range("E11")
WS5.Range("H20") = WS1.Range("E12")
WS5.Range("H21") = WS1.Range("E13")
WS5.Range("H22") = WS1.Range("E14")
WS5.Range("H23") = WS1.Range("E15")
WS5.Range("H24") = WS1.Range("E16")
WS5.Range("H25") = WS1.Range("E17")
WS5.Range("H26") = WS1.Range("E18")

WS5.Range("I15:I26") = WS1.Range("G2") '単位




WS5.Range("J15") = WS1.Range("F7") '単価
WS5.Range("J16") = WS1.Range("F8")
WS5.Range("J17") = WS1.Range("F9")
WS5.Range("J18") = WS1.Range("F10")
WS5.Range("J19") = WS1.Range("F11")
WS5.Range("J20") = WS1.Range("F12")
WS5.Range("J21") = WS1.Range("F13")
WS5.Range("J22") = WS1.Range("F14")
WS5.Range("J23") = WS1.Range("F15")
WS5.Range("J24") = WS1.Range("F16")
WS5.Range("J25") = WS1.Range("F17")
WS5.Range("J26") = WS1.Range("F18")
WS5.Range("J27") = WS1.Range("F19")

'WS5.Range("K15").MergeArea = WS1.Range("G7") '金額
'WS5.Range("K16").MergeArea = WS1.Range("G8")
'WS5.Range("K17").MergeArea = WS1.Range("G9")
'WS5.Range("K18").MergeArea = WS1.Range("G10")
'WS5.Range("K19").MergeArea = WS1.Range("G11")
'WS5.Range("K20").MergeArea = WS1.Range("G12")
'WS5.Range("K21").MergeArea = WS1.Range("G13")
'WS5.Range("K22").MergeArea = WS1.Range("G14")
'WS5.Range("K23").MergeArea = WS1.Range("G15")
'WS5.Range("K24").MergeArea = WS1.Range("G16")
'WS5.Range("K25").MergeArea = WS1.Range("G17")
'WS5.Range("K26").MergeArea = WS1.Range("G18")

Dim FLN As String 'ファイル名
Dim strAa As String '保存
Dim strCa As String
Dim strNa As Object

Set strNa = CreateObject("WScript.NetWork")
strCa = strNa.UserName


FLN = Application.GetSaveAsFilename(RefNum & " " & WS1.Range("C3") & " " & WS1.Range("E4") & ".xlsx", FileFilter:="Excelファイル,*.xlsx")


ActiveWorkbook.SaveAs Filename:=FLN



MsgBox "注番ﾊｲﾌﾝ,単位入力を忘れずに"


'クリア

WS1.Range("C2").ClearContents
WS1.Range("E3:E4").ClearContents
WS1.Range("B7:F19").ClearContents



End Sub

