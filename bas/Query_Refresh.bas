Attribute VB_Name = "Query_Refresh"
Sub Query_Refresh()
Attribute Query_Refresh.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' Query_Refresh Macro
'
' Keyboard Shortcut: Ctrl+p
'
    Range("課題管理[[#Headers],[課題ナンバー]]").Select
    ActiveWorkbook.Connections("クエリ - 課題管理").Refresh
    
    Call kadaicount.kadaicount
    
End Sub
