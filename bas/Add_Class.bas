Attribute VB_Name = "Add_Class"
Sub MakeKamoku()
    Dim Count, ForCt As Integer
    Dim StopSine, Kamoku, TableName As String
    Dim ws As Worksheet
    Dim Flag As Boolean
    
    ForCt = Range("KamokuList[#ALL]").Rows.Count
    For Count = 2 To ForCt Step 1
        
        Worksheets("Debug").Select
        Kamoku = Cells(Count, 1)
        
        Flag = False
        For Each ws In Worksheets
            If ws.Name = Kamoku Then
                Flag = True
            Else

            End If
        Next ws
        
        If Flag = False Then
            'シートをコピーする
            If Not Kamoku = "end" Then
                Worksheets("Template").Copy After:=Worksheets(Worksheets.Count)
                ActiveSheet.Name = Kamoku
                    
                'テーブル名変更
                Range("A2").ListObject.Name = Kamoku
            Else
            End If
                
        ElseIf Flag = True Then
        
        Else
        
        End If

    Next Count
             
End Sub
