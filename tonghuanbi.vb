'@author:fuwenyue
'Created on Sat Apr 1 2023
'
'代码为EXCEL VBA 代码
'用于制作甜啦啦同环比信息表
'

Sub 创建依赖表()
    '创建同比数据源表、环比数据源表、总表、哗啦啦门店信息表
    '门店管理信息表、本期收银源表、环比期收银源表、同比期收银源表、门店类型
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' 创建“同比数据源表”
    Dim tbSrc As Worksheet
    Set tbSrc = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    tbSrc.Name = "同比数据源表"
    
    ' 创建“环比数据源表”
    Dim hbSrc As Worksheet
    Set hbSrc = wb.Sheets.Add(After:=tbSrc)
    hbSrc.Name = "环比数据源表"
    
    ' 创建“总表”
    Dim total As Worksheet
    Set total = wb.Sheets.Add(After:=hbSrc)
    total.Name = "总表"
    
    ' 创建“哗啦啦门店信息表”
    Dim hll As Worksheet
    Set hll = wb.Sheets.Add(After:=total)
    hll.Name = "哗啦啦门店信息表"
    
    ' 创建“门店管理信息表”
    Dim mgmt As Worksheet
    Set mgmt = wb.Sheets.Add(After:=hll)
    mgmt.Name = "门店管理信息表"
    
    ' 创建“本期收银源表”
    Dim currCash As Worksheet
    Set currCash = wb.Sheets.Add(After:=mgmt)
    currCash.Name = "本期收银源表"
    
    ' 创建“环比期收银源表”
    Dim prevCash As Worksheet
    Set prevCash = wb.Sheets.Add(After:=currCash)
    prevCash.Name = "环比期收银源表"
    
    ' 创建“同比期收银源表”
    Dim sameCash As Worksheet
    Set sameCash = wb.Sheets.Add(After:=prevCash)
    sameCash.Name = "同比期收银源表"
    
    ' 创建“门店类型”
    Dim storeType As Worksheet
    Set storeType = wb.Sheets.Add(After:=sameCash)
    storeType.Name = "门店类型"
    
End Sub

Sub 同环比左侧_插入列并填入表头()
    Dim j As Integer
    On Error GoTo Last
    '取消合并A1的单元格
    Range("A1").UnMerge
    '在E列前插入9列
    Columns("E:E").Select
    For j = 1 To 9
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightorAbove
    Next j

    '填入表头
    Range("E3:E4") = "门店ID"
    Range("F3:F4") = "门店编号"
    Range("G3:G4") = "门店类型"
    Range("H3:H4") = "大区经理"
    Range("i3:i4") = "省经理"
    Range("j3:j4") = "区域经理"
    Range("k3:k4") = "运营状态"
    Range("l3:l4") = "是否有收银机"
    Range("m3:m4") = "是否为学校店铺"
Last:     Exit Sub
End Sub

Sub 同环比左侧_填入公式()
    '在左侧插入的列中填入公式
    Dim headers As Variant
    Dim formulas As Variant
    Dim i As Integer
    
    headers = Array("门店ID", "门店编号", "门店类型", "大区经理", "省经理", "区域经理", "运营状态", "是否有收银机", "是否为学校店铺")
    formulas = Array("=XLOOKUP(D5,哗啦啦门店信息表!E:E,哗啦啦门店信息表!D:D)", _
                     "=XLOOKUP(E5,门店管理信息表!M:M,门店管理信息表!B:B)", _
                     "=XLOOKUP(E5,门店类型!D:D,门店类型!AA:AA)", _
                     "=XLOOKUP(F5,门店管理信息表!B:B,门店管理信息表!J:J)", _
                     "=XLOOKUP(F5,门店管理信息表!B:B,门店管理信息表!I:I)", _
                     "=XLOOKUP(F5,门店管理信息表!B:B,门店管理信息表!H:H)", _
                     "=XLOOKUP(F5,门店管理信息表!B:B,门店管理信息表!K:K)", _
                     "=XLOOKUP(F5,门店管理信息表!B:B,门店管理信息表!AR:AR)", _
                     "=XLOOKUP(F5,门店管理信息表!B:B,门店管理信息表!AA:AA)")
    
    '填入公式
    For i = LBound(headers) To UBound(headers)
        Range("E4").Offset(, i).Value = headers(i)
        Range("E5").Offset(, i).Formula = formulas(i)
    Next i
End Sub

Sub 环比右侧_添加列并填入表头()
    Range("AW3:AZ3") = "堂食流水"
    Range("BA3:BD3") = "堂食实收"
    Range("BE3:BH3") = "外卖流水"
    Range("BI3:BL3") = "外卖实收"
    Range("BM3:BP3") = "美团"
    Range("BQ3:BT3") = "饿了么"
    Range("BU3:BX3") = "其它"
    Range("BY3:CB3") = "自提流水"
    Range("CC3:CF3") = "自提实收"
    Range("AW4") = "本期"
    Range("AX4") = "环比期"
    Range("AY4") = "增长"
    Range("AZ4") = "增长%"
    Range("AW4:AZ4").Copy Range("BA4")
    Range("AW4:AZ4").Copy Range("BE4")
    Range("AW4:AZ4").Copy Range("BI4")
    Range("AW4:AZ4").Copy Range("BM4")
    Range("AW4:AZ4").Copy Range("BQ4")
    Range("AW4:AZ4").Copy Range("BU4")
    Range("AW4:AZ4").Copy Range("BY4")
    Range("AW4:AZ4").Copy Range("CC4")
    '合并表头单元格
    Application.DisplayAlerts = False
    Range("AW3:AZ3").Merge
    Range("BA3:BD3").Merge
    Range("BE3:BH3").Merge
    Range("BI3:BL3").Merge
    Range("BM3:BP3").Merge
    Range("BQ3:BT3").Merge
    Range("BU3:BX3").Merge
    Range("BY3:CB3").Merge
    Range("CC3:CF3").Merge
    Application.DisplayAlerts = True
    '修改单元格方便透视
    Range("AW4") = "本期堂食流水"
    Range("AX4") = "环比期堂食流水"
    Range("BA4") = "本期堂食实收"
    Range("BB4") = "环比期堂食实收"
    Range("BE4") = "本期外卖流水"
    Range("BF4") = "环比期外卖流水"
    Range("BI4") = "本期外卖实收"
    Range("BJ4") = "环比期外卖实收"
    Range("BM4") = "本期美团"
    Range("BN4") = "环比期美团"
    Range("BQ4") = "本期饿了么"
    Range("BR4") = "环比期饿了么"
    Range("BU4") = "本期其它"
    Range("BV4") = "环比期其它"
    Range("BY4") = "本期自提流水"
    Range("BZ4") = "环比期自提流水"
    Range("CC4") = "本期自提实收"
    Range("CD4") = "环比期自提实收"

End Sub

Sub 环比右侧_填入公式()
    '堂食流水
    Range("AW5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!M:M,0)"
    Range("AX5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!M:M,0)"
    '堂食实收
    Range("BA5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!N:N,0)"
    Range("BB5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!N:N,0)"
    '外卖流水
    Range("BE5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!P:P,0)"
    Range("BF5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!P:P,0)"
    '外卖实收
    Range("BI5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!Q:Q,0)"
    Range("BJ5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!Q:Q,0)"
    '美团
    Range("BM5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!V:V,0)"
    Range("BN5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!V:V,0)"
    '饿了么
    Range("BQ5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!W:W,0)"
    Range("BR5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!W:W,0)"
    '其它
    Range("BU5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!AA:AA,0)"
    Range("BV5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!AA:AA,0)"
    '自提流水
    Range("BY5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!S:S,0)"
    Range("BZ5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!S:S,0)"
    '自提实收
    Range("CC5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!T:T,0)"
    Range("CD5") = "=XLOOKUP(N5,环比期收银源表!D:D,环比期收银源表!T:T,0)"
    '计算公式
    Range("AY5") = "=AW5-AX5"
    Range("AZ5") = "=AY5/AX5"
    Range("AZ5").NumberFormat = "0.00%"
    Range("AY5:AZ5").Copy Range("BC5")
    Range("AY5:AZ5").Copy Range("BG5")
    Range("AY5:AZ5").Copy Range("BK5")
    Range("AY5:AZ5").Copy Range("BO5")
    Range("AY5:AZ5").Copy Range("BS5")
    Range("AY5:AZ5").Copy Range("BW5")
    Range("AY5:AZ5").Copy Range("CA5")
    Range("AY5:AZ5").Copy Range("CE5")
End Sub

Sub 同比右侧_添加列并填入表头()
    Range("AW3:AZ3") = "堂食流水"
    Range("BA3:BD3") = "堂食实收"
    Range("BE3:BH3") = "外卖流水"
    Range("BI3:BL3") = "外卖实收"
    Range("BM3:BP3") = "美团"
    Range("BQ3:BT3") = "饿了么"
    Range("BU3:BX3") = "其它"
    Range("BY3:CB3") = "自提流水"
    Range("CC3:CF3") = "自提实收"
    Range("AW4") = "本期"
    Range("AX4") = "同比期"
    Range("AY4") = "增长"
    Range("AZ4") = "增长%"
    Range("AW4:AZ4").Copy Range("BA4")
    Range("AW4:AZ4").Copy Range("BE4")
    Range("AW4:AZ4").Copy Range("BI4")
    Range("AW4:AZ4").Copy Range("BM4")
    Range("AW4:AZ4").Copy Range("BQ4")
    Range("AW4:AZ4").Copy Range("BU4")
    Range("AW4:AZ4").Copy Range("BY4")
    Range("AW4:AZ4").Copy Range("CC4")
    '合并表头单元格
    Application.DisplayAlerts = False
    Range("AW3:AZ3").Merge
    Range("BA3:BD3").Merge
    Range("BE3:BH3").Merge
    Range("BI3:BL3").Merge
    Range("BM3:BP3").Merge
    Range("BQ3:BT3").Merge
    Range("BU3:BX3").Merge
    Range("BY3:CB3").Merge
    Range("CC3:CF3").Merge
    Application.DisplayAlerts = True
    '修改单元格方便透视
    Range("AW4") = "本期堂食流水"
    Range("AX4") = "同比期堂食流水"
    Range("BA4") = "本期堂食实收"
    Range("BB4") = "同比期堂食实收"
    Range("BE4") = "本期外卖流水"
    Range("BF4") = "同比期外卖流水"
    Range("BI4") = "本期外卖实收"
    Range("BJ4") = "同比期外卖实收"
    Range("BM4") = "本期美团"
    Range("BN4") = "同比期美团"
    Range("BQ4") = "本期饿了么"
    Range("BR4") = "同比期饿了么"
    Range("BU4") = "本期其它"
    Range("BV4") = "同比期其它"
    Range("BY4") = "本期自提流水"
    Range("BZ4") = "同比期自提流水"
    Range("CC4") = "本期自提实收"
    Range("CD4") = "同比期自提实收"

End Sub

Sub 同比右侧_填入公式()
    '堂食流水
    Range("AW5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!M:M,0)"
    Range("AX5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!M:M,0)"
    '堂食实收
    Range("BA5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!N:N,0)"
    Range("BB5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!N:N,0)"
    '外卖流水
    Range("BE5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!P:P,0)"
    Range("BF5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!P:P,0)"
    '外卖实收
    Range("BI5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!Q:Q,0)"
    Range("BJ5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!Q:Q,0)"
    '美团
    Range("BM5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!V:V,0)"
    Range("BN5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!V:V,0)"
    '饿了么
    Range("BQ5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!W:W,0)"
    Range("BR5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!W:W,0)"
    '其它
    Range("BU5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!AA:AA,0)"
    Range("BV5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!AA:AA,0)"
    '自提流水
    Range("BY5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!S:S,0)"
    Range("BZ5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!S:S,0)"
    '自提实收
    Range("CC5") = "=XLOOKUP(N5,本期收银源表!D:D,本期收银源表!T:T,0)"
    Range("CD5") = "=XLOOKUP(N5,同比期收银源表!D:D,同比期收银源表!T:T,0)"
    '计算公式
    Range("AY5") = "=AW5-AX5"
    Range("AZ5") = "=AY5/AX5"
    Range("AZ5").NumberFormat = "0.00%"
    Range("AY5:AZ5").Copy Range("BC5")
    Range("AY5:AZ5").Copy Range("BG5")
    Range("AY5:AZ5").Copy Range("BK5")
    Range("AY5:AZ5").Copy Range("BO5")
    Range("AY5:AZ5").Copy Range("BS5")
    Range("AY5:AZ5").Copy Range("BW5")
    Range("AY5:AZ5").Copy Range("CA5")
    Range("AY5:AZ5").Copy Range("CE5")
End Sub



Sub 总表_插入新列并填充表头()
    Dim col As Range
    For Each col In Range("E4:CF4").Cells
        If col.Value = "增长%" Then
            '插入两列
            col.Offset(0, 1).EntireColumn.Insert
			col.Offset(0, 1).EntireColumn.ClearFormats
            col.Offset(0, 2).EntireColumn.Insert
            '更新列标题
            col.Offset(0, 1).Value = "同比期"
            col.Offset(0, 2).Value = "同比增长"
        End If
    Next col
    '在PRT列后插入新列
    Columns("U").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("S").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("Q").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Q4") = "同比期"
    Range("T4") = "同比期"
    Range("W4") = "同比期"
End Sub

Sub 总表_填充同比数据()
    
    Dim wb As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim copyRange As Range

    '打开工作簿
    Set wb = ThisWorkbook
    '设置源工作表
    Set ws1 = wb.Worksheets("同比数据源表")
    '设置目标工作表
    Set ws2 = ActiveWorkbook.ActiveSheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, "P").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "S").End(xlUp).Row
    '营业天数
    Set copyRange = ws1.Range("P5:P" & lastRow1)
    copyRange.Copy ws2.Range("Q5")
    '翻台率
    Set copyRange = ws1.Range("R5:R" & lastRow1)
    copyRange.Copy ws2.Range("T5")
    '上座率
    Set copyRange = ws1.Range("T5:T" & lastRow1)
    copyRange.Copy ws2.Range("W5")
    '流水金额
    ws1.Range("V5:V" & lastRow1).Copy ws2.Range("AB5")
    ws1.Range("X5:X" & lastRow1).Copy ws2.Range("AC5")
    '实收金额
    ws1.Range("Z5:Z" & lastRow1).Copy ws2.Range("AH5")
    ws1.Range("AB5:AB" & lastRow1).Copy ws2.Range("AI5")
    '优惠金额
    ws1.Range("AD5:AD" & lastRow1).Copy ws2.Range("AN5")
    ws1.Range("AF5:AF" & lastRow1).Copy ws2.Range("AO5")
    '账单数
    ws1.Range("AH5:AH" & lastRow1).Copy ws2.Range("AT5")
    ws1.Range("AJ5:AJ" & lastRow1).Copy ws2.Range("AU5")
    '客流
    ws1.Range("AL5:AL" & lastRow1).Copy ws2.Range("AZ5")
    ws1.Range("AN5:AN" & lastRow1).Copy ws2.Range("BA5")
    '单均消费
    ws1.Range("AP5:AP" & lastRow1).Copy ws2.Range("BF5")
    ws1.Range("AR5:AR" & lastRow1).Copy ws2.Range("BG5")
    '人均消费
    ws1.Range("AT5:AT" & lastRow1).Copy ws2.Range("BL5")
    ws1.Range("AV5:AV" & lastRow1).Copy ws2.Range("BM5")
    '堂食流水
    ws1.Range("AX5:AX" & lastRow1).Copy
    Sheets("总表").Select
    Range("BR5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("AZ5:AZ" & lastRow1).Copy
    Sheets("总表").Select
    Range("BS5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '堂食实收
    ws1.Range("BB5:BB" & lastRow1).Copy
    Sheets("总表").Select
    Range("BX5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("BC5:BC" & lastRow1).Copy
    Sheets("总表").Select
    Range("BY5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '外卖流水
    ws1.Range("BF5:BF" & lastRow1).Copy
    Sheets("总表").Select
    Range("CD5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("BH5:BH" & lastRow1).Copy
    Sheets("总表").Select
    Range("CE5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '外卖实收
    ws1.Range("BJ5:BJ" & lastRow1).Copy
    Sheets("总表").Select
    Range("CJ5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("BL5:BL" & lastRow1).Copy
    Sheets("总表").Select
    Range("CK5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '美团
    ws1.Range("BN5:BN" & lastRow1).Copy
    Sheets("总表").Select
    Range("CP5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("bp5:bp" & lastRow1).Copy
    Sheets("总表").Select
    Range("CQ5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '饿了么
    ws1.Range("BR5:BR" & lastRow1).Copy
    Sheets("总表").Select
    Range("CV5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("Bt5:Bt" & lastRow1).Copy
    Sheets("总表").Select
    Range("CW5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '其它
    ws1.Range("BV5:BV" & lastRow1).Copy
    Sheets("总表").Select
    Range("DB5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("BX5:BX" & lastRow1).Copy
    Sheets("总表").Select
    Range("DC5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '自提流水
    ws1.Range("BZ5:BZ" & lastRow1).Copy
    Sheets("总表").Select
    Range("DH5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("CB5:CB" & lastRow1).Copy
    Sheets("总表").Select
    Range("DI5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '自提实收
    ws1.Range("CD5:CD" & lastRow1).Copy
    Sheets("总表").Select
    Range("DN5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws1.Range("CF5:CF" & lastRow1).Copy
    Sheets("总表").Select
    Range("DO5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AZ5").NumberFormat = "0.00"
    Range("O4") = "本期营业天数"
    Range("P4") = "环比期营业天数"
    Range("Q4") = "同比期营业天数"	
    Range("AD4") = "本期实收金额"
    Range("AE4") = "环比期实收金额"
    Range("AH4") = "同比期实收金额"	
    Range("BT4") = "本期堂食实收"
    Range("BU4") = "环比期堂食实收"
    Range("BX4") = "同比期堂食实收"	
    Range("CF4") = "本期外卖实收"
    Range("CG4") = "环比期外卖实收"
    Range("CJ4") = "同比期外卖实收"		
    Range("CL4") = "本期美团"
    Range("CM4") = "环比期美团"
    Range("CP4") = "同比期美团"		
    Range("CR4") = "本期饿了么"
    Range("CS4") = "环比期饿了么"
    Range("CV4") = "同比期饿了么"
	Range("CX4") = "本期其它"
    Range("CY4") = "环比期其它"
    Range("DB4") = "同比期其它"	
	Range("DJ4") = "本期自提实收"
    Range("DK4") = "环比期自提实收"
    Range("DN4") = "同比期自提实收"	
End Sub