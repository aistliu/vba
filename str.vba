'把Text放入剪贴板
Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub
'把剪贴板中内容放入到单元格中，如果文本中有换行等，会依次放入单元格的下一个单元格中。
Sub PasteCell()
 Dim a As String 
 a = "222"
 '把a放入剪贴板
 Call CopyText(a)
 '把剪贴板中的内容放到B2中，如果a中有换行，tab，会放入以B2为起始位置的单元格中
 ActiveWorkbook.Worksheets(1).Range("B2").Select
 ActiveWorkbook.Worksheets(1).Paste
End Sub

