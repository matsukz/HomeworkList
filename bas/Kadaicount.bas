Attribute VB_Name = "KadaiCount"
Sub kadaicount()
    Dim Kyou, Ashita, Week, Ato As Integer
    Dim Hikaku As String
    
    Worksheets("�ۑ�Ǘ�").Select
    
    '�����l���Z�b�g
    Kyou = 0: Ashita = 0
    Week = 0: Ato = 0
    
    For i = 3 To Range("�ۑ�Ǘ�[#ALL]").Rows.Count + 1
        If Not Cells(i, 8) = "��o�ς�" Then
            '��o�����܂ł̃X�e�[�^�X���擾
            Hikaku = Cells(i, 5)
            Select Case Hikaku
                Case "��t�I��"
                    '��t�I�����Ă���̂ŃJ�E���g���Ȃ�
                Case "����"
                    Kyou = Kyou + 1
                Case "����1��"
                    Ashita = Ashita + 1
                Case Else
                    '�c������2���ȏ�̂Ƃ�
                    '�c��������i�[����ϐ���ς���
                    If Cells(i, 4) - Date <= 7 Then
                        Week = Week + 1
                    Else
                        Ato = Ato + 1
                    End If
                End Select
        Else
        End If
    Next i
    
    '�Z���ɃZ�b�g
    Worksheets("�ۑ�o�^").Select
    
    Range("G4") = Kyou
    Range("G5") = Ashita
    Range("G6") = Week
    Range("G7") = Ato
    
    Worksheets("�ۑ�Ǘ�").Select
End Sub
        

