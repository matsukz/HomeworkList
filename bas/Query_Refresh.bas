Attribute VB_Name = "Query_Refresh"
Sub Query_Refresh()
Attribute Query_Refresh.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' Query_Refresh Macro
'
' Keyboard Shortcut: Ctrl+p
'
    Range("�ۑ�Ǘ�[[#Headers],[�ۑ�i���o�[]]").Select
    ActiveWorkbook.Connections("�N�G�� - �ۑ�Ǘ�").Refresh
End Sub
