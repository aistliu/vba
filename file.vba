'显示打开
Sub workbook_operate()
 ' 定义工作薄对象
    Dim wbk As Workbook
    Dim fname As String
    
    fname = "E:/temp/test.xlsx"
    ' 根据工作薄文件路径打开工作薄
    Set wbk = Application.Workbooks.Open(Filename:=fname)‘类似的还有字典等变量，别忘了用set
    MsgBox fname & "已打开"
    ' 关闭工作薄 这种方式会弹出对话框
    wbk.Close 
    ‘wbk.Close SaveChanges:=False ’不显示是否保存文件的对话框直接关闭 False不保存关闭，True保存关闭
End Sub

'隐式打开，电脑上不会看到文件打开了
Sub workbook_operate()
    ' 定义工作薄对象
    Dim wbk As Workbook
    Dim fname As String
    
    fname = "E:/temp/test.xlsx"
    ' 根据工作薄文件路径获取工作薄对象
    Set wbk = GetObject(fname)
    Debug.Print wbk.Name
End Sub

’ThisWorkbook与ActiveWorkbook
Sub workbook_operate()

    ' 定义工作薄对象
    Dim wbk As Workbook
    Dim fname As String
    
    fname = "E:/temp/ActiveMe.xlsx"
    ' 根据工作薄文件路径获取工作薄对象
    Set wbk = Workbooks.Open(fname)
  
    Debug.Print ThisWorkbook.Name
    Debug.Print ActiveWorkbook.Name
End Sub
'打印出文件夹下所有文件名字
Sub OpenAndClose()
    Dim MyFile As String
    Dim s As String
    Dim count As Integer
    MyFile = Dir("C:\Users\McDelfino\Desktop\2.JPL_SCAT_EXCEL全\" & "*.xlsx")
    '读入文件夹中的第一个.xlsx文件
    count = count + 1       '记录文件的个数
    s = s & count & "、" & MyFile
    Do While MyFile <> ""
        MyFile = Dir        '第二次读入的时候不用写参数
        If MyFile = "" Then
            Exit Do         '当MyFile为空的时候就说明已经遍历完了，这时退出Do，否则还要运行一遍
        End If
        count = count + 1
        If count Mod 2 <> 1 Then
            s = s & vbTab & count & "、" & MyFile
        Else
            s = s & vbCrLf & count & "、" & MyFile
        End If
    Loop
    Debug.Print s
End Sub
