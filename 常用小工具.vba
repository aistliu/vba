
'一个单元格中有多行，把N行拆成N个并放入N个row，其他列信息原样拷贝。
Sub chaihange()
    Dim maxR As Integer, firstR As Integer, firstC As Integer
    firstR As ActiveCell.Row
    firstC As ActiveCell.Column
    maxR = ActiveSheet.Cells(Rows.Count, firstC).End(xlUp).Row
    
    Dim r As Long
    Dim tblsIN As String
    Dim tblArr() As String
    Dim tblUbound As Integer
    
    Dim addr As Integer
    For r = maxR To firstR Step -1
        tblsIN = Trim(ActiveSheet.Cells(r, firstC))
        
        If InStr(tblsN, Chr(10)) > 0 Then
            tblArr = slit(tblsIN, Chr(10))
            tblUbound = UBound(tblArr)
            
            ActiveSheet.Rows(r & ":" & r).Select
            Selection.Copy
            ActiveSheet.Rows(r + 1 & ":" & r + UBound(tblArr)).Select
            Selection.Insert shift:=xlDown
            For addr = r To r + tblUbound
                ActiveSheet.Cells(addr, firstC) = tblArr(addr - r)
            Next
        End If
    Next
End Sub
