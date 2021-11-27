
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
'在筛选后的列中，按照当前的行顺序拷贝进去，新拷贝进入的数据以当前行为准
Sub copyToFilterCol()
    On Error Resume Next
    '需要以下的参照设定（工具->引用）Microsoft Forms 2.0 Object Library｡如果没有需要以下文件 c:\Windows\System32\FM20.DLL
    Dim dataObj As New MSForms.DataObject
    dataObj.GetFromClipboard
    kw = dataObj.GetText
    
    Dim clipArr() As String
    clipArr = Split(kw, vbCrLf)
    
    Dim rng As Range
    Dim i As Integer
    i = 0
    
    Dim offsetRow As Integer
    offsetRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    For Each rng In ActiveSheet.Range(ActiveCell, ActiveCell.Offset(100000, 0)).SpecialCells(xlCellTypeVisible)
        rng = Replace(clipArr(i), """", "")
        i = i + 1
        If i > UBound(clipArr) Then
            Exit For
        End If
    Next rng
End Sub
'当前sheet中的某个关键字，全部变红
Sub changePartColerInShtRed()
    On Error Resume Next
    Dim kw As strng
    Dim dataObj As New MSForms.DataObject
    dataObj.GetFromClipboard
    kw = dataObj.GetText
    Call changepartColerInSht(kw, 3)'1黑2白3红4绿5蓝6黄
    kw = ""
End Sub
Private Sub changepartColerInSht(kw As String, colorIdx As Integer)
    Dim r As Range
    Dim cellV As String
    Dim pos As Integer
    
    For Each r In ActiveSheet.UsedRange
        cellV = r.Value
        
        pos = 1
        kwLen Len(kw)
        
        Do While True
            pos = InStr(pos, cellV, kw)
            If pos = 0 Then
                Exit Do
            Else
                r.Characters(Start:=pos, Length:=kwLen).Font.ColorIndex = colorIdx
            End If
        Loop
    Next r
End Sub

