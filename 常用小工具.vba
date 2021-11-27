
'一个单元格中有N行时（用Chr（10）换行符分割的字符串），此行下面插入N-1个空行，其他列信息原样拷贝，次列信息
'按照换行符拆成N个独立字符串分别写入此列
Sub chaihange()
    Dim maxR As Integer, firstR As Integer, firstC As Integer
    firstR = ActiveCell.Row
    firstC = ActiveCell.Column
    maxR = ActiveSheet.Cells(Rows.Count, firstC).End(xlUp).Row
    
    Dim r As Long
    Dim tmpCellVal As String
    Dim arrByChr10() As String
    Dim arrUbound As Integer
    
    Dim addr As Integer
    For r = maxR To firstR Step -1
        tmpCellVal = Trim(ActiveSheet.Cells(r, firstC))
        
        If InStr(tmpCellVal, Chr(10)) > 0 Then
            arrByChr10 = Split(tmpCellVal, Chr(10))
            arrUbound = UBound(arrByChr10)
            
            ActiveSheet.Rows(r & ":" & r).Select
            Selection.Copy
            ActiveSheet.Rows(r + 1 & ":" & r + UBound(arrByChr10)).Select
            Selection.Insert shift:=xlDown
            For addr = r To r + arrUbound
                ActiveSheet.Cells(addr, firstC) = arrByChr10(addr - r)
            Next
        End If
    Next
End Sub

