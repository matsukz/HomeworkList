Attribute VB_Name = "Registration"
Sub Touroku()

    Worksheets("�ۑ�o�^").Select

    Dim Kamoku, Hyodai, Bikou, TableName As String
    Dim InstCount As Integer
    Dim St_Date, En_Date As Date
    
    '�e�ϐ��ւ̊i�[
    Kamoku = Range("C3")
    Hyodai = Range("C4")
    St_Date = Range("C5")
    En_Date = Range("C6")
    Bikou = Range("C7")
    
    
    '���[�N�V�[�g�ړ�
    Worksheets(Kamoku).Select
    '�e�[�u���I��
    TableName = Kamoku & "[#ALL]"
    InstCount = Range(TableName).Rows.Count + 2
    

    '�]�L�J�n
    '�ۑ�i���o�[
    Cells(InstCount, 1) = Cells(InstCount - 1, 1) + 1
    
    '�\��
    Cells(InstCount, 2) = Hyodai
    
    '�e����t
    Cells(InstCount, 3) = St_Date
    Cells(InstCount, 4) = En_Date
    
    '�[��/���t�̓V�[�g��Ōv�Z����
    
    '�����K��
    Cells(InstCount, 6) = Kamoku
    
    '�i��/��o
    Cells(InstCount, 7) = "������"
    Cells(InstCount, 8) = "����o"
    
    '���l
    Cells(InstCount, 9) = Bikou
    
    '��n��
    Worksheets("�ۑ�Ǘ�").Select
    ActiveWorkbook.Connections("�N�G�� - �ۑ�Ǘ�").Refresh
    
    Worksheets("�ۑ�o�^").Select
    Range("C3").Select
    
End Sub
