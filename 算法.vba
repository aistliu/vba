Sub digui(n As Integer, arr() As String)
    If n >= 1 Then
        arr(UBound(arr)) = n
        n = n - 1
        If n >= 1 Then
            ReDim Preserve arr(UBound(arr) + 1)
        End If
        Call digui(n, arr)
    End If
End Sub
Sub testDigui()
    Dim n As Integer, arr() As String
    n = 5
    ReDim arr(0)
    Call digui(n, arr)
    Dim i
    For Each i In arr
        Debug.Print i
    Next i
End Sub
