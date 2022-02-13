'sleep lib
Private Declare Sub Sleep Lib "kernel32" (ByVal ms as Long )
'数组和range的转换，2维数组直接拷贝到range，速度很快 
Private Sub arrRange()
    Dim arr() As String
    ReDim arr(3, 2)
    
    '2 wei shu zu
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            arr(i, j) = i & "," & j
        Next
    Next
    
    Dim sh As Worksheet
    Set sh = ActiveSheet
    Dim ra As Range
    Set re = sh.Range(sh.Cells(2, 1).Address, sh.Cells(5, 3).Address)
    re.Value = arr
    
    Dim newArr
    
    newArr = sh.UsedRange
End Sub
'字典中存放（数组，collection，字典）的写法
Private Sub addObjToDic()
    Dim d As New Dictionary
    Dim c As New Collection
    ' add array to dic
    d.add "1", Array("a", "b")
    'add collection to dic
    c.add "item", "key1"
    c.add "item2", "key2"
    d.add "coll", c
    'add dic to dic
    d.add "dic1", New Dictionary
    d("dic1").add "keyDic", "valDic"
End Sub

'跳过各种对话框，直接打开文件，加快执行速度，关闭在计算，屏幕刷新，更新链接对话框等等
Private Sub speedUpInXls()
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.ThisWorkbook.UpdateLinks = xlUpdateLinksNever
    
    'slip readOlny dialoge
    filePath "c:\aaa.xls"
    Workbooks.Open FileName:=filePath, Password:="", ReadOnly:=readOlnyFlg, IgnoreReadonlyRecommended:=True
    
    Application.ThisWorkbook.UpdateLinks = xlUpdateLinksAlways
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
End Sub
'拷贝粘贴的写法
Private Sub sheet_copy_paste()
    Dim srcWb As Workbook, desWb As Workbook
    Dim srcPath As String, desPath As String
    srcPath = "c:\s.xls"
    desPath = "c:\d.xls"
    
    'copy sheet to  d.xls
    Set srcWb = Lib.getWB(srcPath, True)
    Set desWb = Lib.getWB(srcPath, False)
    srcWb.Worksheets("Sheet1").Copy before:=desWb.Worksheets(1)
    
    'copy shape to sheet
    Windows("s.xls").Activate
    ActiveSheet.Shapes(1).Select
    Selection.Copy
    Windows("d.xls").Activate
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveSheet.Paste
End Sub
'插入行列
Private Sub insert_Rows_columns()
    Dim sh As Worksheet
    Set sh = ActiveSheet
    
    sh.Rows(3 & ":" & 5).Insert  ' After inserted,  new lines:3,4,5 old lines: 1,2,6,7..., line6 is old line3
    'insert column
    sh.Columns(3).Insert
    sh.Columns("A").Insert
    sh.Range("B2").EntireColumn.Insert shift:=xlToRight
    sh.Range("B:C").Insert
    sh.Range(Columns(2), Columns(5)).Insert
    
    sh.Range(sh.Cells(2, 2), sh.Cells(3, 3)).Insert shift:=xlToRight
End Sub
'先过滤行，然后循环过滤后的行
Private Sub loopFilteredRows(sh As Worksheet)
    With sh.UsedRange
        .AutoFilter Field:=1, Criteria1:="maru", Operator:=xlOr, Criteria2:="sannkaku"
        .AutoFilter Field:=2, Criteria1:="2maru"
    End With
    Dim r As Long
    Dim rng As Range
    For Each rng In sh.Range("A2:A" & sh.UsedRange.SpecialCells(xlCellTypeLastCell).Row).SpecialCells(xlCellTypeVisible)
        r = rng.Row
        ' debug.Pring sh.cells(r,1).Value
    Next rng
End Sub
'使用excel自带的函数
Private Sub useXlsFunction()
    Dim b As Integer
    With Application.WorksheetFunction
        b = .Match("AA*", ActiveWorkbook.Worksheets("Sheet1").Range(Cells(1, 1), Cells(12, 1)), 0)
    End With
    Debug.Print b
End Sub








