
Private Declare Sub Sleep Lib "kernel32" (ByVal ms as Long )'sleep lib
Private Const SPSTR As String = "@@#$*@@"
'workbook,sheet-------------------------->
Function maxR(sh As Worksheet, col As Integer)
    maxR = sh.Cells(Rows.Count, col).End(xlUp).Row
End Function
Function maxC(sh As Worksheet, r As Integer)
    maxC = sh.Cells(r, Columns.Count).End(xlToLeft).Column
End Function
'get workbook
Function getWB(filePath, Optional readOnlyFlg As Boolean = False) As Workbook
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    Dim wb1 As Workbook
    Dim FileName As String
    On Error Resume Next
    FileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    Set wb1 = Workbooks(FileName) ' if xls file is opened,use this opened file only by xls Name
    'if not opened ,error 9 will occur, then open the xls by full path
    If Err.Number = 9 Then
        Workbooks.Open FileName:=filePath, Password:="", ReadOnly:=readOlnyFlg, IgnoreReadonlyRecommended:=True
        Set wb1 = ActiveWorkbook
    End If
    
    Set getWB = wb1
    
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Function

Sub copyFromTemplateXls(templateFilePath, newFileNamePath)
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    Workbooks.add (templateFilePath)
    ActiveWorkbook.SaveAs newFileNamePath
    ActiveWorkbook.Close saveChanges:=True
    
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub

Function shtHasKeyword(sht As Worksheet, kw As String) As Boolean
    shtHasKeyword = False
    Dim r As Range
    With sht.UsedRange
        Set r = .Cells.Find(What:=kw)
        If Not r Is Nothing Then
            shtHasKeyword = True
            Exit For
        End If
    End With
End Function
Sub deleteFilterRows(sht As Worksheet, colCo As Integer, filterKW As String)
    With sht.UsedRange
        .AutoFilter Field:=colCo, criterial1:=filterKW
    End With
    sht.UsedRange.Offset(1, 0).Resize(sht.UsedRange.Rows.Count - 1).Rows.Delete
End Sub
'da xiao xie tong yiaBc,aBc,ABC, AbC,ABc..=> aBc
Sub unityStrInSht(sht As Worksheet)
    Dim i, j, arr, tmp
    Dim dic As New Dictionary
    arr = sht.UsedRange
    
    For i = LBound(arr, 1) To UBound(arr, 1)
            
        For j = LBound(arr, 2) To UBound(arr, 2)
            tmp = Trim(arr(i, j))
            If Not dic.Exists(UCase(tmp)) Then
                dic.add UCase(tmp), tmp
                arr(i, j) = tmp
            Else
                arr(i, j) = dic(UCase(tmp))
            End If
        Next
    Next
    sht.UsedRange.Value = arr
    
End Sub
'workbook,sheet--------------------------<

'Array-------------------------->
Function arrIsNull(arr) As Boolean
    arrIsNull = False
    On Error Resume Next
    If (UBound(arr)) < 0 Then
        arrIsNull = True
    End If
End Function
Function arrExists(arr, item) As Boolean
    arrExists = False
    Dim i
    On Error GoTo ErrorHangler
    For Each i In arr
        If i = item Then
            arrExists = True
            Exit For
        End If
    Next i
ErrorHandler:
    If Not arrExists Then
        arrExists = False
    End If
End Function
Sub arrRemove(arr, item)
    arr = Filter(arr, item, False)
End Sub
Sub arrAdd(arr, item)
    If arrIsNull(arr) Then
        ReDim arr(0)
    Else
        ' not exist, then add
        If Not arrExists(arr, item) Then
            ReDim Preserve arr(UBound(arr) + 1)
            arr(UBound(arr)) = item
        End If
End Sub
Function arrMerge(arr1, arr2)
    If arrIsNull(arr1) Then
        arrMerge = arr2
    Else
        arrMerge = Split(Join(arr1, SPSTR) & SPSTR & Join(arr2, SPSTR), SPSTR)
End Function
Function arrRemoveDuplicate()
    Dim d As New Dictionary
    For Each i In arr
        If Not d.Exists(i) Then
            d.add i, ""
        End If
    Next i
    
    Dim ret() As String
    ReDim ret(d.Count - 1)
    Dim ind As Long
    ind = 0
    For Each k In d
        ret(ind) = k
        ind = ind + 1
    Next k
    arrRemoveDuplicate = ret
End Function
'Array--------------------------<
'File--------------------------------->
Private Function getFso() As Object
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set getFso = fso
End Function
Sub FileCopy(srcPath, desPath)
    FileCopy srcPath, desPath
End Sub
Sub getFilesList(folderPath As String, searchSub As Boolean, arr() As String)
    Dim fname
    fname = Dir(folderPath & "\*.*")
    
    On Error Resume Next
    If UBound(arr) < 1 Then
        ReDim arr(0)
    End If
    
    Do While fname <> ""
        arr(UBound(arr)) = folderPath & "\" & fname
        If fname <> "" Then
            ReDim Preserve arr(UBound(arr) + 1)
        End If
        fname = Dir
    Loop
    
    If searchSub Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject").GetFolder(folderPath)
        Dim folder
        For Each folder In fso.SubFolders
            If InStr(folder, "Arxhive") <= 0 Then
                Call getFilesList(folder.path, searchSub, arr)
            Else
                Debug.Print " Aechive folder:" & folder
            End If
        Nextfolder
    End If
End Sub
Function getFileExt(filePath As String)
    Dim fso As Object
    Set fso = getFso()
    getFileExt = fso.GetExtensionName(filePath)
End Function
Function FileExists(filePath As String) As Boolean
    Dim str As String
    str = ""
    On Error Resume Next
    str = Dir(filePath)
    On Error GoTo 0
    If str = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function
Function FileExists2(filePath As String) As Boolean
    Dim fso As Object
    Set fso = getFso()
    With fso
        If .FileExists(filePath) Then
            FileExists2 = True
        Else
            FileExists2 = False
        End If
    End With
End Function
Function txtLineCount(fp As String) As Long
    Dim fso As Object
    Set fso = getFso
    txtLineCount = fso.OpenTextFile(FileName:=fp, IOMode:=8).Line
    Set fso = Nothing
End Function
Function fileNameFromPath(path As String)
    Dim ret As String
    If Len(path) > 0 Then
        ret = Right(path, Len(path) - InStrRev(path, "\"))
        ret = Left(ret, InStrRev(ret, ".") - 1)
        fileNameFromPath = ret
    End If
End Function
Function fileLastModifiedTime(f)
    Dim fso As Object
    Set fso = getFso
    fileLastModifiedTime = fso.GetFile(f).DateLasModified
End Function
Function fileTxt(path As String)
    Dim cs As String
    cs = fncGetCharset(path)
    Dim buffer As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    
    If cs = "UTF-8" Or cs = "Unicode" Or cs = "Shift_JIS" Then
        stream.Charset = cs
    Else
        stream.Charset = "_autodetect_all"
    End If
    
    stream.LoadFromFile (path)
    buffer = stream.ReadText
    stream.Close
    fileTxt = buffer
    Set stream = Nothing
End Function
Function fileTxtByLine(path As String)
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
Sub wirteTxt(path As String, str As String)
    Open path For Output As #2
    Print #2, str
    Print #2, "other word...."
    Close #2
End Sub
Function fncGetCharset(FileName As String) As String
    Dim i                   As Long     '?用指数
       
    Dim lngFileLen          As Long     'ファイルサイズ
    Dim bytFile()           As Byte     'ファイル内容
    Dim b1                  As Byte     '1バイト目
    Dim b2                  As Byte     '2バイト目
    Dim b3                  As Byte     '3バイト目
    Dim b4                  As Byte     '4バイト目
       
    Dim lngSJIS             As Long     'Shift_JISの可能性
    Dim lngUTF8             As Long     'UTF-8もの可能性
    Dim lngEUC              As Long     'EUC-JPの可能性
     
    'ADODB定数
    Const adModeUnknown = 0
    Const adModeRead = 1
    Const adModeWrite = 2
    Const adModeReadWrite = 3
    Const adModeShareDenyRead = 4
    Const adModeShareDenyWrite = 8
    Const adModeShareExclusive = 12
    Const adModeShareDenyNone = 16
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    Const adReadLine = -2
     
    'ファイル?み?み（バイナリ?）
    On Error Resume Next
    With CreateObject("ADODB.Stream")
        .Mode = adModeUnknown
        .Open
        .Type = adTypeBinary
        .LoadFromFile FileName
        lngFileLen = .Size
        bytFile = .read(adReadAll)
        .Close
    End With
    If (Err.Number <> 0) Then
        fncGetCharset = "OPEN FAILED"
        Exit Function
    End If
    On Error GoTo 0
     
    'BOMによる判断
    If (bytFile(0) = &HEF And bytFile(1) = &HBB And bytFile(2) = &HBF) Then
        fncGetCharset = "UTF-8 BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFF And bytFile(1) = &HFE) Then
        fncGetCharset = "UTF-16 LE BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFE And bytFile(1) = &HFF) Then
        fncGetCharset = "UTF-16 BE BOM"
        Exit Function
    End If
       
    'BINARY
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If ((b1 >= &H0 And b1 <= &H1F) And b1 <> &H9 And b1 <> &HA And b1 <> &HD And b1 <> &H1B) Or (b1 = &H7F) Then
            fncGetCharset = "BINARY"
            Exit Function
        End If
    Next i
              
    'SJIS
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Or (b1 >= &HB0 And b1 <= &HDF) Then
            lngSJIS = lngSJIS + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                   ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   lngSJIS = lngSJIS + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
              
    'UTF-8
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngUTF8 = lngUTF8 + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   lngUTF8 = lngUTF8 + 2
                   i = i + 1
                Else
                    If (i < lngFileLen - 3) Then
                        b3 = bytFile(i + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) Then
                            lngUTF8 = lngUTF8 + 3
                            i = i + 2
                        Else
                            If (i < lngFileLen - 4) Then
                                b4 = bytFile(i + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) And (b4 >= &H80 And b4 <= &HBF) Then
                                    lngUTF8 = lngUTF8 + 4
                                    i = i + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
   
    'EUC-JP
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngEUC = lngEUC + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And _
                   (b2 >= &HA1 And b2 <= &HFE)) Or _
                   ((b1 = &H8E) And (b2 >= &HA1 And b2 <= &HDF)) Then
                   lngEUC = lngEUC + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
              
    '文字コ?ド出??位による判断
    If (lngSJIS <= lngUTF8) And (lngEUC <= lngUTF8) Then
        fncGetCharset = "UTF-8"
        Exit Function
    End If
    If (lngUTF8 <= lngSJIS) And (lngEUC <= lngSJIS) Then
        fncGetCharset = "Shift_JIS"
        Exit Function
    End If
    If (lngUTF8 <= lngEUC) And (lngSJIS <= lngEUC) Then
        fncGetCharset = "EUC-JP"
        Exit Function
    End If
     
    '判定不能
    fncGetCharset = "UNKNOWN"
End Function

Sub waitFileUntilFileIsCreate(filePath As String)
    Dim hasFile As Boolean
    hasFile = False
    On Error Resume Next
    Do Until hasFile
        Sleep 2000
        hasFile = FileExists(filePath)
    Loop
    'Kill filePath ' if need, delete this file
End Sub

'File---------------------------------<
'regexp------------------------------------------------------------------------->
Function regExists(src As String, re As String) As Boolean
    regExists = False
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    With reg
        .Pattern = re
        .ignorecase = True
        .Global = True
    End With
    
    If ret.test(src) Then
        regExists = True
    End If
End Function

Function regFind(src As String, re As String) As String()
    Dim ret() As Strin
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    With reg
        .Pattern = re
        .ignorecase = True
        .Global = True
    End With
    
    Dim matches As Object
    Dim matche As Object
    Set matches = reg.Execute(src)
    If matches.Count > 0 Then
        ReDim ret(matches.Count - 1)
        Dim i As Integer
        For i = 0 To matches.Count - 1
            ret(i) = matches(i)
        Next
    End If
    regFind = ret
End Function

Function regReplace(sr As String, oldStrReg As String, newStr As String) As String
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    With reg
        .Pattern = oldStrReg
        .ignorecase = True
        .Global = True
    End With
    ret = reg.Replace(src, newStr)
    replaceReg = CStr(ret)
End Function
'regexp-------------------------------------------------------------------------<
'shell---------------------------------------------->
' run xxx.bat
Private Sub exeShell(batFile As String)
    Dim argh As Double
    argh = Shell(batFile, vbNornalFocus)
End Sub
Private Sub openTxtFile(filePath As String)
    Shell "cmd /c start """" explorer.exe " & filePath, vbHide
End Sub
'shell---------------------------------------------->

'windows----------------------------------------------------------->
' put txt to clipboard
Sub SetStrToClipboard(txt As String)
    Dim msobj
    Set msobj = CreateObject("new:{IC3B4210-F441-11CE-B9EA-00AA006B1A69}")
    msobj.SetText txt
    msobj.PutInClipboard
End Sub
'get txt from clipboard
Function getStrFromClipboard() As String
    Dim dataObj As New MSForms.DataObject
    dataObj.GetFromClipboard
    getStrFromClipboard = dataObj.GetText
End Function
'windows-----------------------------------------------------------<
