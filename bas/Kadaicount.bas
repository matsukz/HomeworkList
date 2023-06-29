Attribute VB_Name = "KadaiCount"
Sub kadaicount()
    Dim Kyou, Ashita, Week, Ato As Integer
    Dim Hikaku As String
    
    Worksheets("課題管理").Select
    
    '初期値をセット
    Kyou = 0: Ashita = 0
    Week = 0: Ato = 0
    
    For i = 3 To Range("課題管理[#ALL]").Rows.Count + 1
        If Not Cells(i, 8) = "提出済み" Then
            '提出期限までのステータスを取得
            Hikaku = Cells(i, 5)
            Select Case Hikaku
                Case "受付終了"
                    '受付終了しているのでカウントしない
                Case "今日"
                    Kyou = Kyou + 1
                Case "あと1日"
                    Ashita = Ashita + 1
                Case Else
                    '残日数が2日以上のとき
                    '残日数から格納する変数を変える
                    If Cells(i, 4) - Date <= 7 Then
                        Week = Week + 1
                    Else
                        Ato = Ato + 1
                    End If
                End Select
        Else
        End If
    Next i
    
    'セルにセット
    Worksheets("課題登録").Select
    
    Range("G4") = Kyou
    Range("G5") = Ashita
    Range("G6") = Week
    Range("G7") = Ato
    
    Worksheets("課題管理").Select
End Sub
        

