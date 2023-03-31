Sub 筛选省经理()
    ActiveSheet.Range("$P$3:$P$173").AutoFilter Field:=1, Criteria1:=RGB(255, _
        192, 0), Operator:=xlFilterCellColor, VisibleDropDown:=False
End Sub
Sub 筛选大区经理()
    ActiveSheet.Range("$P$3:$P$173").AutoFilter Field:=1, Criteria1:=RGB(255, _
        255, 0), Operator:=xlFilterCellColor, VisibleDropDown:=False
End Sub
Sub 筛选区域经理()
    ActiveSheet.Range("$P$3:$P$173").AutoFilter Field:=1, Operator:= _
        xlFilterNoFill, VisibleDropDown:=False
End Sub
Sub 取消筛选()
    ActiveSheet.Range("$S$3:$S$173").AutoFilter Field:=1
End Sub