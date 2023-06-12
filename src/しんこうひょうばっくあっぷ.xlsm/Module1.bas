Attribute VB_Name = "Module1"

Sub �敪����1(dummy)

Dim LastRow As Long
Dim i As Long
    
    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '�ŏI�s�܂Ń��[�v
    For i = 17 To LastRow
    
        '�X�ԍ����g����O�����
        On Error Resume Next
        
        With Cells(i, "D")
            .Offset(0, 11) = WorksheetFunction.VLookup(.Value, Worksheets("Sheet1").Range("A:C"), 3, False)
    
        End With
        
        '�c�Ə����g����P�����
        With Cells(i, "E")
            .Offset(0, 11) = WorksheetFunction.VLookup(.Value, Worksheets("Sheet1").Range("B:C"), 3, False)
    
        End With
        
        On Error GoTo 0
        
    Next

End Sub

Sub �敪����2(dummy)

Dim LastRow As Long
Dim i As Long

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '�ŏI�s�܂Ń��[�v
        For i = 17 To LastRow

            '����̑��t��̎��̋敪����
            If InStr(Cells(i, "E"), "���эH�|��") >= 1 Then
                .Cells(i, "B") = "�ڊǕ�"
                
            ElseIf InStr(Cells(i, "E"), "DNP���W�X�e�B�N�X") >= 1 Then
                .Cells(i, "B") = "�]����"
                
            End If
            
        Next
        
    End With
    
End Sub

Sub �񐿋��폜(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim Target As Variant

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '�Ώە���
    Target = Array("�񐿋�")

    Application.DisplayAlerts = False
        With ActiveSheet

            '�ŏI�s����擪�փ��[�v
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target)

                    '���蕶������������s�폜
                    If InStr(Cells(i, 13), Target(j)) >= 1 Then
                        .Rows(i).Delete
                    End If

                Next
                
            Next
            
        End With
    Application.DisplayAlerts = True

End Sub

Sub �R�[�q�[�J�^���O�u��(dummy)

Dim LastRow As Long
Dim i As Long

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '�ŏI�s����擪�փ��[�v
        For i = LastRow To 1 Step -1

            On Error Resume Next
            
            '�R�[�q�[�ƃJ�^���O�u��
            If InStr(Cells(i, "F"), "�R�[�q�[") >= 1 Then
                .Cells(i, "F").ClearContents
                .Cells(i, "F") = "FF�R�[�q�[�֘A"
                
            ElseIf InStr(Cells(i, "F"), "FAMIMACAFE") >= 1 Then
                .Cells(i, "F").ClearContents
                .Cells(i, "F") = "FF�R�[�q�[�֘A"
                
            ElseIf InStr(Cells(i, "F"), "�J�^���O") >= 1 Then
                .Cells(i, "F").ClearContents
                .Cells(i, "F") = "�J�^���O�֘A"
                
            End If
            
            On Error GoTo 0
            
        Next
        
    End With
    
End Sub

Sub �ʑ��ǉ�(dummy)

Dim LastRow As Long
Dim myDate As Date '���s��
Dim fstDate As Date '����
Dim i As Long
Dim bs As Borders
Dim myMon As Date

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    myDate = Range("L5")
    fstDate = Year(myDate) & "/" & Month(myDate) & "/" & 1
    
    i = GetWeekDay(fstDate, Range("L5"), vbTuesday)

    '�s�̒ǉ�
    Rows(LastRow + 1).Insert
    
    '�w�i�F�����F�ɕύX
    Range(Cells(LastRow + 1, "A"), Cells(LastRow + 1, "L")).Interior.ColorIndex = 6
    
    '�g���̒ǉ�
    Set bs = Range(Cells(LastRow + 1, "A"), Cells(LastRow + 1, "L")).Borders
    bs.LineStyle = xlContinuous
    
    '�����̒ǉ�
    Cells(LastRow + 1, "E") = "�R�`��c�Ə�"
    Cells(LastRow + 1, "F") = "FXSS�V���[�J�[�h�ʑ�"
    Cells(LastRow + 1, "G") = 1
    Cells(LastRow + 1, "H") = 185
    Cells(LastRow + 1, "I") = 185
    Cells(LastRow + 1, "J") = 600
    Cells(LastRow + 1, "K") = 600
    Cells(LastRow + 1, "L") = 785
    
    '�s�̕���
    Rows(LastRow + 1).Copy
    Range(Rows(LastRow + i - 1), Rows(LastRow + 1)).Insert
    
End Sub

Sub ���t���\�[�g(dummy)

Dim LastRow As Long
Dim sht As Worksheet

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '�\�[�g�̃N���A
        .Sort.SortFields.Clear
    
        '�\�[�g�̐ݒ�
        .Sort.SortFields.Add _
            Key:=ActiveSheet.Cells(16, "A"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        
        '���t�����Ƀ\�[�g
        With .Sort
            .SetRange Range(Cells(16, "A"), Cells(LastRow - 1, "L"))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
        End With
        
    End With

    Set sht = ActiveSheet
    sht.Cells.Interior.ColorIndex = xlNone

End Sub

Sub �G���A���P���ݒ�(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim Target1 As Variant
Dim Target2 As Variant
Dim Target3 As Variant
    
    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '�Ώە���
    Target1 = Array("���l��", "��B1", "��B2", "���B")
    Target2 = Array("����")
    Target3 = Array("�����{")
    
    Application.DisplayAlerts = False

        With ActiveSheet

            '�ŏI�s����擪�փ��[�v
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target1)
                For k = 0 To UBound(Target2)
                For l = 0 To UBound(Target3)

                    'Target1
                    If InStr(Cells(i, "O"), Target1(j)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If
                    
                    If InStr(Cells(i, "P"), Target1(j)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If
                    
                    If InStr(Cells(i, "Q"), Target1(j)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If
                    
                    'Target2
                    If InStr(Cells(i, "O"), Target2(k)) >= 1 Then
                        .Cells(i, "J").ClearContents
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "P"), Target2(k)) >= 1 Then
                        .Cells(i, "J").ClearContents
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "Q"), Target2(k)) >= 1 Then
                        .Cells(i, "J").ClearContents
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                
                    'Target3
                    If InStr(Cells(i, "Q"), Target3(l)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If

                Next l
                Next k
                Next j
            
            Next i
        
        End With
    
    Application.DisplayAlerts = True

End Sub

Sub �P���F�t��(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim Target1 As Variant
Dim Target2 As Variant
Dim Target3 As Variant
Dim Target4 As Variant

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '�Ώە���
    Target1 = Array("���}�g", "���}�g�^�C���T�[�r�X")
    Target2 = Array("DM")
    Target3 = Array("�q���")
    Target4 = Array("�Z�[���i�̂ݑ��t", "���ꎮ", "��̑����ꎮ")
    
    Application.DisplayAlerts = False
    
        With ActiveSheet

            '�ŏI�s����擪�փ��[�v
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target1)
                For k = 0 To UBound(Target2)
                For l = 0 To UBound(Target3)
                For m = 0 To UBound(Target4)

                    '���蕶������������Z���̐F�ύX
                    If InStr(Cells(i, "N"), Target1(j)) >= 1 Then
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "N"), Target2(k)) >= 1 Then
                        .Cells(i, "J").Interior.ColorIndex = xlNone
                    End If
                    
                    If InStr(Cells(i, "M"), Target3(l)) >= 1 Then
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "M"), Target4(m)) >= 1 Then
                        .Cells(i, "G").Interior.ColorIndex = 6
                        .Cells(i, "H").Interior.ColorIndex = 6
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If

                Next m
                Next l
                Next k
                Next j
                
            Next
            
        End With
        
    Application.DisplayAlerts = True

End Sub

Sub ���ꎮ�ɕ����ǉ�(dummy)

Dim LastRow As Long
Dim i As Long

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '�ŏI�s����擪�փ��[�v
        For i = LastRow To 1 Step -1

            On Error Resume Next
            
                '"���ꎮ"������Δ��l�ɕ����ǉ�
                If InStr(Cells(i, "M"), "���ꎮ") >= 1 Then
                    Cells(i, "M") = Cells(i, "M").Value & "�^�����̂ݏT����ƂƓ���P��"
                End If
            
            On Error GoTo 0
            
        Next
        
    End With
    
End Sub

Sub �����Ɣ�ύX(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim Target As Variant

    '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '�Ώە���
    Target = Array("���ꎮ", "�Z�[���i�̂ݑ��t")

    Application.DisplayAlerts = False
        With ActiveSheet

            '�ŏI�s����擪�փ��[�v
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target)

                    '���蕶�����������獫���Ɣ��ύX
                    If InStr(Cells(i, "M"), Target(j)) >= 1 Then
                        .Cells(i, "H") = 24
                    End If

                Next
                
            Next
            
        End With
    Application.DisplayAlerts = True

End Sub

Sub �V�[�g�����ƕs�v�s�폜(dummy)

Dim LastRow1 As Long
Dim LastRow2 As Long
Dim i As Long
Dim j As Long
Dim targetSheet As Worksheet

    '�V�[�g�̖��O�ύX
    ActiveSheet.Name = "�ǉ�����_DM��"
    
    '�V�[�g�̕���
    Worksheets("�ǉ�����_DM��").Copy Before:=Worksheets(1)
    
    '���������V�[�g�̖��O�ύX
    ActiveSheet.Name = "�ǉ�����_��z��"

    '�ŏI�s�擾
    LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row

        '�ŏI�s����擪�փ��[�v
        For i = LastRow1 To 1 Step -1
        
            On Error Resume Next

                '"���}�gDM��"�̍s���폜
                If InStr(Cells(i, "N"), "���}�gDM��") >= 1 Then
                    Rows(i).Delete
                End If
                
            On Error GoTo 0
                    
        Next
        
    '�s�v����폜
    Range("N:Q").Delete
            
    '�A�N�e�B�u�V�[�g��"�ǉ�����_DM��"�ɕύX
    Sheets("�ǉ�����_DM��").Select
    
    '�ŏI�s�擾
    LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

        '�ŏI�s����擪�փ��[�v
        For j = LastRow2 To 1 Step -1
        
            On Error Resume Next

                '"���}�gDM��"�ȊO�̍s���폜
                If InStr(Cells(j, "N"), "���}�gDM��") = 0 Then
                    Rows(j).Delete
                End If
                
            On Error GoTo 0
                    
        Next
        
    '�s�v����폜
    Range("N:Q").Delete
    
    '�����Ă��܂����w�b�_�[���̍s��}��
    Range("1:16").Insert
        
    '���}�gDM�ֈȊO�����������ŏ������^�C�g�����̃R�s�y
    Sheets("�ǉ�����_��z��").Select
    Range("1:16").Copy Sheets("�ǉ�����_DM��").Range("A1")
    
    
    '�s�v���[�N�V�[�g�폜
    Application.DisplayAlerts = False
    
    For Each targetSheet In Worksheets
    
        If Not targetSheet.Name Like "�ǉ�����*" Then
            targetSheet.Delete
            
        End If
    
    Next
    
    Application.DisplayAlerts = True
            
End Sub


