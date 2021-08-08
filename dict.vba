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
  
  Msgbox dic("a") ' 显示key为a的值：苹果
End Sub
