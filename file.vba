'递归取得文件夹下所有文件路径名list，存入传入的数组
Private Sub getFolderFilesList(pth As String, doSubFolder As Boolean, arr() As String)
    Dim fName
    fName = Dir(pth & "\*.*") '获取文件下的第一个文件名（不含路径），不限扩展名
    
    Do While fName <> ""
        arr(UBound(arr)) = pth & "\" & fName
        If fName <> "" Then
            ReDim Preserve arr(UBound(arr) + 1)
        End If
        fName = Dir '获取文件下的次一个文件名
    Loop
    
'递归查找子文件夹
    If doSubFolder Then
       Dim fso, folder
    'fso主要方法，fso.path, fso.files/获取该路径下所有文件名子文件夹名　fso.FileExsts(path)/文件存在否，返回true/false
       Set fso = CreateObject("Scripting.FileSystemObject").getFolder(pth)
       For Each folder In fso.SubFlders
           Call getFolderFilesList(folder.Path, doSubFolder, arr)’递归子文件夹
       Next folder
    Next
End Sub

'读入文件路径名，返回该文件的文本---＞
Function fileTxt(path As String)
    Dim encode As String
    encode = GetEncoding(path)
    
    If encode = "Unicode" Then
        fileTxt = fileTxtUn(path)
    Else
        fileTxt = fileTxtByLine(path)
    End If
End Function
Private Function fileTxtUn(path As String) As String
    Dim buffer As String, stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "Unicode" 'UTF-8/SHIFT-JIS
    stream.Type = 2 '2表示是text
    stream.Open
    stream.LoafFromFile (path)
    buffer = stream.ReadText
    stream.Close
End Function
Private Function fileTxtByLine(path As String) As stirng
    Dim txt As String
    Dim tmpTxt As String
    If Dir(path) = "" Then
        fileTxtByLine = ""
    Else
        Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpTxt
            txt = txt & tmpTxt & vbCrLf
        Loop
        Close #1
        fileTxtByLine = txt
    End If
End Function
Private Function GetEncoding(fileName)
    Dim fBytes(1) As Byte
    Open fileName For Binary Access Read As #1
    Get #1, , fBytes(0)
    Get #1, , fBytes(1)
    Close #1
    GetEncoding = IIf(fBytes(0) = &HFF And fBytes(1) = &HFE, "Unicode", "ANSI")
End Function
'读入文件路径名，返回该文件的文本---＜

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
