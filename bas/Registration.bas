Attribute VB_Name = "Registration"
Sub Touroku()

    Worksheets("課題登録").Select

    Dim Kamoku, Hyodai, Bikou, TableName As String
    Dim InstCount As Integer
    Dim St_Date, En_Date As Date
    
    '各変数への格納
    Kamoku = Range("C3")
    Hyodai = Range("C4")
    St_Date = Range("C5")
    En_Date = Range("C6")
    Bikou = Range("C7")
    
    
    'ワークシート移動
    Worksheets(Kamoku).Select
    'テーブル選択
    TableName = Kamoku & "[#ALL]"
    InstCount = Range(TableName).Rows.Count + 2
    

    '転記開始
    '課題ナンバー
    Cells(InstCount, 1) = Cells(InstCount - 1, 1) + 1
    
    '表題
    Cells(InstCount, 2) = Hyodai
    
    '各種日付
    Cells(InstCount, 3) = St_Date
    Cells(InstCount, 4) = En_Date
    
    '納期/日付はシート上で計算する
    
    '命名規則
    Cells(InstCount, 6) = Kamoku
    
    '進捗/提出
    Cells(InstCount, 7) = "未完成"
    Cells(InstCount, 8) = "未提出"
    
    '備考
    Cells(InstCount, 9) = Bikou
    
    '後始末
    Worksheets("課題管理").Select
    ActiveWorkbook.Connections("クエリ - 課題管理").Refresh
    
    Worksheets("課題登録").Select
    Range("C3").Select
    
End Sub
