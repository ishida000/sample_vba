Sub 複数シート選択()
    Sheets(Array("Sheet1", "Sheet2", "Sheet3")).Select
End Sub

Sub 複数シート解除()
    ActiveWindow.SelectedSheets(1).Select
End Sub