'字符串src中使用re正则，取得所有的命中项，组成数组返回
Function searchReg(src As String, re As String) As String()
    Dim ret() As String
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    With ret
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
            ret (i) - matches(i)
        Next
    End If
    searchReg = ret
End Function
'正则替换字符串
Function replaceReg(src As String, oldStr As String, newStr As String) As String
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    With ret
        .Pattern = oldStr
        .ignorecase = True
        .Global = True
    End With
    ret = ret.Replace(src, newStr)
    replaceReg = CStr(ret)
End Function
'正则判断src中存在
Function containsReg(src As String, re As String) As Boolean
    containsReg = False
    Dim reg As Object
    Set reg = CreateObject("vbscript.regexp")
    With ret
        .Pattern = oldStr
        .ignorecase = True
        .Global = True
    End With
    If reg.test(src) Then
        containsReg = True
    End If
End Function
'---------------------测试用---------------
Sub Sample1()
    Dim RE, strPattern As String, r As Range
    Set RE = CreateObject("VBScript.RegExp")
    strPattern = "SUM\("
    With RE
        .Pattern = strPattern       ''検索パターンを設定
        .IgnoreCase = True          ''大文字と小文字を区別しない
        .Global = True              ''文字列全体を検索
        For Each r In ActiveSheet.UsedRange
            If .Test(r.Formula) Then r.Interior.ColorIndex = 3
        Next r
    End With
    Set RE = Nothing
End Sub

Sub Sample2()
    Dim RE, strPattern As String, i As Long, msg As String, reMatch
    Set RE = CreateObject("VBScript.RegExp")
    strPattern = "^田(中|口).*(子|美)$"
    With RE
        .Pattern = strPattern
        .IgnoreCase = True
        .Global = True
        For i = 1 To 10
            Set reMatch = .Execute(Cells(i, 1))
            If reMatch.Count > 0 Then
                msg = msg & reMatch(0).Value & vbCrLf
            End If
        Next i
    End With
    MsgBox msg
    Set reMatch = Nothing
    Set RE = Nothing
End Sub

