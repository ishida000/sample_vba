Sub 例外処理()
On Error GoTo ErrorHandler

ErrorHandler:
    'エラー処理
    If Err.Number <> 0 Then
        MsgBox "エラー番号：" & Err.Number & vbCrLf & "エラー内容：" & Err.Description
    End If
End Sub