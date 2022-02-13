Sub month()
    dim mm as integer
    
    '今月の月を取得
    mm1 = Trim(Month(Date))
    '翌月の月を取得
    mm2 = Trim(Month(DateAdd("m", 1, Date)))
    '翌々月の月を取得
    mm3 = Trim(Month(DateAdd("m", 2, Date)))

    Print mm1
    Print mm2
    Print mm3
End Sub


