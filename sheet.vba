Sub 遍历工作表()
  For Each sh In ActiveWorkbooks.Worksheets    '数组
  If sh.Name Like "*" & "表" & "*" Then     '如果工作表名称包含“表”,则选中并弹出对话框显示表名
      sh.Select
      Msgbox sh.Name
    End if
  Next
End Sub
'快速遍历工作簿查询字符串，返回【sheet名：地址】的数组
Function searchWB(wb As Workbook, kw As String) As String()
    Dim loopSht As Worksheet'工作簿要遍历的工作表
    Dim searchRng As Range’工作表中遍历到的单元格
    Dim rngAddr As String‘单元格地址
    Dim addrStr As String’保存地址用的字符串，用逗号分割，分割后返回数组， 保存形式为 工作表名：地址
    
    For Each loopSht In wb.Worksheets
        With loopSht.UsedRange
            Set searchRng = .Cells.Find(What:=kw)‘快速查询工作表中字符串的关键步骤Cells.Find
            If Not searchRng Is Nothing Then
                Dim addr
                addr = searchRng.Address
                Do
                    If Len(addr) > 0 Then
                        addrStr = addrStr & "," & loopSht.Name & ":" & searchRng.Address(ReferenceStyle:=xlA1)  'xlR1C1
                    Else
                        addrStr = loopSht.Name & ":" & searchRng.Address(ReferenceStyle:=xlA1)
                    End If
                Loop While searchRng.Address <> addr
            End If
            
        End With
        If InStr(addrStr, ",") > 0 Then
            searchWB = Split(addrStr, ",")
        Else
            searchWB = Null
        End If
    Next loopSht
End Function
  
'按行循环工作表中的内容
Sub loopSheetLines ()
  Dim sht As Worksheet
  Dim lineIndex As  Integer
  Dim str As String
  Dim bhColor As Integer' 单元格的背景色值，可以据此判断循环的终止等
  Dim StandColor As Integer’基准色，某特定单元格的背景色，用来作为判断的基准色值
  Set sht = ActiveWorkbook.Worksheets(“sheet名”) ‘或者用数字值做参数，表示第几个shhet
  StandColor =  sht.Range("A1").Interior.ColorIndex '得到当前单元格背景色作为循环中的判断标准
  
  dim lastRow as Integer
  lastRow = lastRowByCol(sht, 3) 'C列最后一行行号
  For lineIdx = 3 to lastRowByCol ‘知道循环开始结束行号的时候的循环
    bgColor = sht.Range("A" & lineIdx).Interior.ColorIndex '得到当前单元格背景色
    If bgColor = StandColor Then
      '单元格背景色是某值的时候，
    End If
      
    
    str = Trim(sht.Range("A" & lineIdx).Value)' 得到A列当前行的值
  Next
  '第二种Do while循环
  Do while True
    '循环终止条件，当取得的值是空的时候
    If Len(str) < 1 Then
      Exit Do
    End If
    str = Trim(sht.Range("A" & lineIdx).Value)' 得到A列当前行的值
    
    If InStr(str, "AABB“) >0 Then
      '当字符串中含有AABB的时候的操作
    End If
    
    lineIdx = lineIdx + 1 '循环的行加一
  Loop
End Sub
'某列的最终行
Private Function lastRowByCol(ByRef sht As Worksheet, colNum As Integer) As Integer
    Dim xlLastRow As Long
    Dim LastRow As Long
 
    xlLastRow = sht.Cells(Rows.Count, 1).Row
    LastRow = sht.Cells(xlLastRow, colNum).End(xlUp).Row
    lastRowByCol = LastRow
End Function
