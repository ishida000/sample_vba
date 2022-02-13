Sub Mailbox_Open()
    Call メール作成(customer_name)
End Sub

Function メール作成(ship_name As Variant)
    '--- Outlook操作のオブジェクト ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    '--- メールオブジェクト ---'
    Dim objMail As Object
    Set objMail = objOutlook.CreateItem(0)
        
    '--- メールの内容を格納する変数 ---'
    Dim toStr As String
    Dim ccStr As String
    Dim bccStr As String
    Dim subjectStr As String
    Dim bodyStr As String
    
    '--- 宛先の内容 ---'
    toStr = Worksheets("更新日").Range("G3").Value '"[宛先のメールアドレス]"
    ccStr = "" '"[CCのメールアドレス]"
    bccStr = "" ' "[BCCのメールアドレス]"
    
    '--- 件名の内容 ---'
    subjectStr = "販売計画更新通知_" + ship_name
    
    '--- 本文の内容 ---'
    bodyStr = "お疲れ様です。 " + vbCrLf + vbCrLf + "各自対応をお願いします｡"
        
    '--- 条件を設定 ---'
    objMail.To = toStr
    objMail.CC = ccStr
    objMail.BCC = bccStr
    objMail.Subject = subjectStr
    'objMail.BodyFormat = olFormatPlain
    objMail.Body = bodyStr
    
    '--- 添付ファイルのパス ---'
    Dim attachmentPath As String
    attachmentPath = "\\Afnewt320-kyoyu\社内共有\【生産管理】\販売計画\" + ship_name + ".xlsm"

    
    '--- 添付ファイルを設定 ---'
    Call objMail.Attachments.Add(attachmentPath)
    
    '--- メールを表示 ---'
    objMail.Display
    
    '--- メールを送付 ---'
    'objMail.Send

End Function