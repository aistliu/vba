https://www.cnblogs.com/dazuo/p/4920921.html
https://for200x.pixnet.net/blog/post/210814375-excel-vba-%E6%8C%87%E4%BB%A4%E9%9B%86%E9%8C%A6
VBA中字符換行顯示需要使用換行符來完成。下面是常用的換行符 
            chr(10)可以生成換行符 
            chr(13)可以生成ENTER 
            vbcrlf換行符和ENTER 
            vbCr等同於chr(10) 
            vblf等同於chr(13) 
  

·  EXCEL常用的物件

Workbook 活頁簿
Workbooks 活頁簿集合
Workbooks("filename") 檔名為filename的活頁簿
ActiveWorkbook 正在作用中的活頁簿
Sheets 活頁簿中所有工作表
Sheets(n) 活頁簿中第n張工作表
Worksheet 工作表
Worksheets 所有工作表(包括圖表)
Worksheets("sheet") 指表名為sheet工作表
ActiveSheet 正在作用中的工作表
Columns("c1:c2") c1至c2欄(其中c1,c2為A~Z或AA~XFD等欄名)
Rows("r1:r2") r1至r2列(其中r1,r2為1~1048576等列名
Range("x1:x2") x1至x2間的儲存格(其中x1,x2為儲存格位址名稱)
cells(i,j) 儲存格(第i列、第j行)
ActiveCell 目前的儲存格
Selection 目前所選取的物件
·  範例：

Workbooks("Book1").Sheets("Sheet1").Range("A1:D5").Font.Bold = True
Worksheets("Sheet1").Cells.ClearContents
Worksheets("Sheet1").Rows(1).Font.Bold = True
Range("1:1,3:3,8:8")
Worksheets("Sheet1").Cells(6, 1).Value = 10
Worksheets("Sheet1").[A1:B5].ClearContents
ActiveCell.Offset(1, 3).Font.Underline = xlDouble
·  活頁簿常用屬性：

ActiveWorkBook.Name 目前活頁簿的名稱
ActiveWorkBook.Save 儲存目前的活頁簿
ActiveWorkBook.SaveAs Filename := "filename" 另儲新檔
WorkBooks.Add 新增活頁簿
WorkBooks(i).Close [SaveChange, Filename, RouteWorkbook] 關閉指定的第i個活頁簿
SaveChange := True 改變儲存
SaveChange := False 不會改變儲存
SaveChange省略時，會出現對話方塊
filename := "檔名"
WorkBooks.Open "filename" 開啟一個活頁簿
Application.Windows 所有活頁簿視窗
WorkBooks.Count 活頁簿的數量
WorkBooks.Item(Index) 傳回單一活頁簿，由索引值指定
·  工作表常用屬性：

Worksheets.Add [Before, After, Count, Type] 新增工作表
Before := Worksheets(n) 出現於某工作表之前
After := Worksheets(n) 出現於某工作表之後
Count := n 新增工作表數量
Type := xlWorksheet (工作表) 或 xlChart (圖表)
WorkSheets.Name 工作表名稱
WorkSheets("Sheet1").Activate 設定工作表為目前作用的功作表
·  儲存格常用屬性：

Rows.RowHeight 指定範圍內的所有列高
Columns.ColumnsWidth：指定範圍內的所欄寬
expression.NumberFormatLocal 以本地的數字格式
Range.CurrentRegion 目前區域是指以任意空白列及空白欄的組合為邊界的範圍
範例：
Worksheets("Sheet1").Activate
ActiveCell.CurrentRegion.Select

 



 

expression.Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) 以參照的方式
RowAbsolute 為True，則用列的絕對位址
ColumnAbsolute 為True，則用欄的絕對位址
ReferenceStyle 預設值為xlA1，如為xlR1C1則為R1C1的表達方式
 

expression.count 傳回範圍的數量(可以是欄數、列數或儲存格數量)
expression.Item(RowIndex, ColumnIndex) 代表相對於指定之範圍某個位移距離的範圍。
expression.value 傳回或設定物件的值
expression.Formula 傳回或設定物件的公式，代表 A1 樣式註解以及巨集語言中的物件公式。

範例：Worksheets("Sheet1").Range("A1").Formula = "=$A$4+$A$10"
expression.FormulaR1C1 傳回或設定物件的公式，並以巨集語言中的 R1C1 樣式標記法表示

範例：Worksheets("Sheet1").Range("B1").FormulaR1C1 = "=SQRT(R1C1)"
expression.Text 傳回或設定物件的文字
 
範例：
   Set c = Worksheets("Sheet1").Range("B14")
   c.Value = 1198.3
   c.NumberFormat = "$#,##0_);($#,##0)"
   MsgBox c.Value
   MsgBox c.Text

 


·  常用方法：
 

Range.Select方法/Selection屬性 設定目前選取的範圍/使用目前所選取的範圍
 
範例：
Sub Macro1()
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Address"
    Range("A1:B1").Select
    Selection.Font.Bold = True
End Sub

 

 

 

expression.Copy 將目前所選取的物件複製至剪貼簿
expression.Cut 將目前所選取的物件剪下
expression.Delete 將目前所選取的物件刪除
expression.Paste 將剪貼簿的內容貼上
 
範例：
Sub CopyRow()
    Worksheets("Sheet1").Rows(1).Copy
    Worksheets("Sheet2").Select
    Worksheets("Sheet2").Rows(1).Select
    Worksheets("Sheet2").Paste
End Sub

 

expression.RasteSpecial(Paste,Operation, SkipBlanks, Transpose)
 
範例：
With Worksheets("Sheet1")
    .Range("C1:C5").Copy
    .Range("D1:D5").PasteSpecial _
        Operation:=xlPasteSpecialOperationAdd
End With

 

Range.Activate 目前的儲存格
Range.Clear 清除資料
Range.ClearContents 清除資料內容
Range.ClearFormats 清除資料格式
Range.ClearComments 清除註解
expression.AutoFit：自動調整列高和欄寬
Range.FillDown、Range.FillUp、Range.FillLeft、Range.FillRight 填滿
Range.Offset (RowOffset, ColumnOffset) 指定區域的位移列與行
 
範例：
Sub MoveActive()
    Worksheets("Sheet1").Activate
    Range("A1:D10").Select
    ActiveCell.Value = "Monthly Totals"
    ActiveCell.Offset(0, 1).Activate
End Sub

程式語法：
Dim 陳述式(變數)
Dim varname [ As [New] type]

type 包括 Byte、Boolean、Integer、Long、Single、Double、Date、String、Object等

 

 
Set 陳述式(物件)
Set objectvar = {[New] objectexpression | Nothing}
例：Set RangeA = Range("A1:B2")
 

範例：
Sub Random()
    Dim myRange As Range
    Set myRange = Worksheets("Sheet1").Range("A1:D5")
    myRange.Formula = "=RAND()"
    myRange.Font.Bold = True
End Sub

 

 

With 多種屬性設定
 
With 物件

.屬性1 = 設定值

.屬性2 = 設定值

.... End With
 

範例：
Sub AddNew()
Set NewBook = Workbooks.Add
    With NewBook
        .Title = "All Sales"
        .Subject = "Sales"
        .SaveAs Filename:="Allsales.xls"
    End With
End Sub

Array 陣列
Array(Range1, Range2, ....)
 

範例：
Sub Several()
    Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Select
End Sub

InputBox 函數
InputBox("文字說明",[,title][,default][,xpos][,ypos][,helpfile, context])

MsgBox 函數
MsgBox "文字說明"

Union 將多個範圍合併成單一Range物件
Union(Range1, Range2, ...)
 

範例：
Sub MultipleRange()
    Dim r1, r2, myMultipleRange As Range
    Set r1 = Sheets("Sheet1").Range("A1:B2")
    Set r2 = Sheets("Sheet1").Range("C3:D4")
    Set myMultipleRange = Union(r1, r2)
    myMultipleRange.Font.Bold = True
End Sub

 

 

For... Next 陳述式​​​​​​​​​​​​​​
For counter = start to end [ step stepvalue]

[statements]

[Exit For]

[statements]

Next [counter]

範例：
Sub CycleThrough()
    Dim Counter As Integer
    For Counter = 1 To 20
        Worksheets("Sheet1").Cells(Counter, 3).Value = Counter
    Next Counter
End Sub

For Each... Next 陳述式
For Each element In group

[statements]

[Exit For]

[statements]

Next [element]
 

範例：
Sub ApplyColor()
    Const Limit As Integer = 25
    For Each c In Range("MyRange")
        If c.Value > Limit Then
            c.Interior.ColorIndex = 27
        End If
    Next c
End Sub

 

 

Do ... Loop 陳述式
Do [{While | Until} condition]

[statements]

[Exit Do]

[statements]

Loop


或

Do

[statements]

[Exit Do]

[statements]

Loop [{While | Until} condition]

If ... Then ... Else ... 陳述式
If condition Then [statements][Else elsestatements]
或
If condition Then

[statements]

[ElseIf condition-n Then

[elseifstatements]...

[Else
 

 

 

[elsestatements]]

End If
