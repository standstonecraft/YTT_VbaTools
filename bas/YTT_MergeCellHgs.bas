Attribute VB_Name = "YTT_MergeCellHgs"
'''
''' ���ᎆ��ŕ\�����ۂȂǂɎg�p�B
''' �I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�
'''
Sub YTT_���ᎆ�Z������_������()
Attribute YTT_���ᎆ�Z������_������.VB_Description = "���ᎆ��ŕ\�����ۂȂǂɎg�p�B\r\n�I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�"
    YTT_���ᎆ�Z������ True, False
End Sub
'''
''' ���ᎆ��ŕ\�����ۂȂǂɎg�p
''' �I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�
'''
Sub YTT_���ᎆ�Z������_�r��()
Attribute YTT_���ᎆ�Z������_�r��.VB_Description = "���ᎆ��ŕ\�����ۂȂǂɎg�p�B\r\n�I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�"
    YTT_���ᎆ�Z������ False, True
End Sub
'''
''' ���ᎆ��ŕ\�����ۂȂǂɎg�p
''' �I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�
'''
Sub YTT_���ᎆ�Z������_������_�r��()
Attribute YTT_���ᎆ�Z������_������_�r��.VB_Description = "���ᎆ��ŕ\�����ۂȂǂɎg�p�B\r\n�I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�"
    YTT_���ᎆ�Z������ True, True
End Sub
'''
''' ���ᎆ��ŕ\�����ۂȂǂɎg�p
''' �I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�
'''
Sub YTT_���ᎆ�Z������(centering As Boolean, surrounding As Boolean)
Attribute YTT_���ᎆ�Z������.VB_Description = "���ᎆ��ŕ\�����ۂȂǂɎg�p�B\r\n�I��͈͓��ŕ����������Ă���񂩂玟�ɕ����������Ă����̎�O�܂ł��������邱�Ƃ��J��Ԃ�"
    
    If Selection.Count > 300 Then
        MsgBox "�͈͂��傫�����܂��i>300�j", vbCritical
        Exit Sub
    End If
    If Selection.Rows.Count > 4 Then
        MsgBox "�s�����傫�����܂��i>4�j", vbCritical
        Exit Sub
    End If

    '���݂̃V�[�g
    Dim sh As Worksheet
    Set sh = ActiveSheet
    '�^�[�Q�b�g�s
    Dim thisRow As Integer
    thisRow = Selection(1).Row
    '�J�n��A�|�C���^��
    Dim pointerCol As Integer
    pointerCol = Selection(1).Column
    '�I����
    Dim endCol As Integer
    endCol = Selection.Column + Selection.Columns.Count
    '�s��
    Dim rowSize As Integer
    rowSize = Selection.Rows.Count
    
    
    '�܂���������
    Selection.MergeCells = False
    
    '�|�C���^�[���ŏI����z����܂ŌJ��Ԃ�
    While pointerCol < endCol
        '�|�C���^��̒l�̌�
        Dim baseValCnt As Integer
        baseValCnt = valCount(sh.Cells(thisRow, pointerCol).Resize(rowSize))
        
        '��������񐔂𒲂ׂ�
        Dim colSize As Integer
        For colSize = 2 To 100
            '�ŏI����z����
            If pointerCol + colSize - 1 >= endCol Then
                Exit For
            End If
            '�V�����͈͂̒l�̌�
            Dim rngValCnt As Integer
            rngValCnt = valCount(sh.Cells(thisRow, pointerCol).Resize(rowSize, colSize))
            '���̗�ɐi������
            If baseValCnt < rngValCnt Then
                Exit For
            End If
        Next colSize
        '���̗�܂��͍ŏI�̎��̗�ɂȂ��Ă���̂ň���炷
        colSize = colSize - 1
        
        '��������Z��
        Dim rngToMerge As range
        Set rngToMerge = sh.Cells(thisRow, pointerCol).Resize(rowSize, colSize)
        
        '��������
        rngToMerge.MergeCells = True
        
        '������
        If centering Then
            rngToMerge.VerticalAlignment = xlCenter
            rngToMerge.HorizontalAlignment = xlCenter
        End If
        
        '�r���ň͂�
        If surrounding Then
            rngToMerge.Borders.LineStyle = xlContinuous
        End If
        
        '�|�C���^��i�߂�
        pointerCol = pointerCol + colSize
    Wend
End Sub

'''
''' �͈͓��̒l�̌���Ԃ�
'''
Private Function valCount(rng As range) As Integer
    valCount = WorksheetFunction.CountIf(rng, "<>")
End Function

