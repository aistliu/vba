'数组移除元素 Filter\true，除此元素外全部删除，Filter\false，这个元素以外全部删除
'Array("a","b") 数组构造函数
’Redim Preserve arr(Ubound(arr)+1) 保持原数组内容扩容一位
‘for each i in arr i不能定义成String只能定义为 Dim i
'数组移除元素 
Private Sub remove(arr, item)
    If Not IsEmpty(arr) Then
        arr = Filter(arr, item, False)
    End If
End Sub
'数组追加元素
Private Sub add(arr, item)
    If IsEmpty(arr) Then
        arr = Array(item)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = item
    End If
End Sub
’两个数组合并
Private Function merge(arr1, arr2)
    merge = Split(Join(arr1, ",") & "," & Join(arr2, ","), ",")
End Function
‘数组中元素存在判断
Private Function exists(arr, item) As Boolean
    exites = False
    Dim i
    For Each i In arr
        If i = item Then
            exists = True
        End If
    Next i
End Function
’追加一个匿名字典的方式
Private Sub addSubDicToDic()
Dim dic As New Dictionary ‘字典声明方式1
    dic.add "subDic1", New Dictionary
    dic("subDic1").add "subDicKey", "subDicValue"
End Sub
'字典声明方式2
Sub dic()
  Dim dic
  Set dic = CreateObject("Scripting.Dictionary")
  dic.Add "a", "苹果"
  dic.Add "b", "香蕉"
  dic.Add "c", "雪梨"
  '存在判断
  If dic.Exists("a") Then
    'key存在判断
  End If

End Sub
