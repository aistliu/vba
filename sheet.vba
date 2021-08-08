Sub 遍历工作表()
  For Each sh In Worksheets    '数组
    If sh.Name Like "*" & "表" & "*" Then     '如果工作表名称包含“表”
    sh.Select
    Msgbox sh.Name
  Next
End Sub
'按行循环工作表中的内容
Sub loopSheetLines ()
  Dim sht As Worksheet
  Dim lineIndex As  Integer
  Dim str As String
  Dim bhColor As Integer' 单元格的背景色值，可以据此判断循环的终止等
  Dim StandColor As Integer’基准色，某特定单元格的背景色，用来作为判断的基准色值
  Set sht = ActiveWorkbook.Worksheets(“sheet名”) ‘或者用数字值做参数，表示第几个shhet
  StandColor =  sht.Range("A1").Interior.ColorIndex '得到当前单元格背景色作为循环中的判断标准

  For lineIdx = 3 to 100 ‘知道循环开始结束行号的时候的循环
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
