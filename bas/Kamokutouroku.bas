Attribute VB_Name = "Kamokutouroku"
Sub MakeKamoku()
    Dim Count As Integer
    Dim StopSine, Kamoku, TableName As String
    
    Count = 0
    While Not StopSine = ""
    
        '�V�[�g���R�s�[����
        Worksheets("Template").Copy After:=Worksheets(Worksheets.Count)
        Kamoku = Cells(2 + Count, 1)
        ActiveSheet.Name (Kamoku)
        
        '�e�[�u�����ύX
        TableName = Range("A2").ListObject.Name
        TableName.Name = Kamoku
        
End Sub
