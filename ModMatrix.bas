Attribute VB_Name = "ModMatrix"
Option Explicit
'�s����g�����v�Z
'��֊֐�
Function �t�s��(Hairetu)
    �t�s�� = F_Minverse(Hairetu)
End Function
Function �s��(Hairetu)
    �s�� = F_MDeterm(Hairetu)
End Function


Function F_MMult(ByVal Hairetu1, ByVal Hairetu2)
    'F_MMult(Hairetu1, Hairetu2)
    'F_MMult(�z��@,�z��A)
    '�s��̐ς��v�Z
    '20180213����
    '20210603����
    
    '���͒l�̃`�F�b�N�ƏC��������������������������������������������������������
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(Hairetu1, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(Hairetu2, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����1�Ȃ玟��2�ɂ���B��)�z��(1 to N)���z��(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Hairetu1 = Application.Transpose(Hairetu1)
    End If
    If IsEmpty(JigenCheck2) Then
        Hairetu2 = Application.Transpose(Hairetu2)
    End If
    
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If UBound(Hairetu1, 1) = 0 Or UBound(Hairetu1, 2) = 0 Then
        Hairetu1 = Application.Transpose(Application.Transpose(Hairetu1))
    End If
    If UBound(Hairetu2, 1) = 0 Or UBound(Hairetu2, 2) = 0 Then
        Hairetu2 = Application.Transpose(Application.Transpose(Hairetu2))
    End If
    
    '���͒l�̃`�F�b�N
    If UBound(Hairetu1, 2) <> UBound(Hairetu2, 1) Then
        MsgBox ("�z��1�̗񐔂Ɣz��2�̍s������v���܂���B" & vbLf & _
               "(�o��) = (�z��1)(�z��2)")
        Stop
        End
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim M2%
    Dim Output#() '�o�͂���z��
    N = UBound(Hairetu1, 1) '�z��1�̍s��
    M = UBound(Hairetu1, 2) '�z��1�̗�
    M2 = UBound(Hairetu2, 2) '�z��2�̗�
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '�e�s
        For J = 1 To M2 '�e��
            For K = 1 To M '(�z��1��I�s)��(�z��2��J��)���|�����킹��
                Output(I, J) = Output(I, J) + Hairetu1(I, K) * Hairetu2(K, J)
            Next K
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_MMult = Output
    
    End Function
Sub �����s�񂩃`�F�b�N(Hairetu)
    '20210603�ǉ�
    
    If UBound(Hairetu, 1) <> UBound(Hairetu, 2) Then
        MsgBox ("�����s�����͂��Ă�������" & vbLf & _
                "���͂��ꂽ�z��̗v�f����" & "�u" & _
                UBound(Hairetu, 1) & "�~" & UBound(Hairetu, 2) & "�v" & "�ł�")
        Stop
        End
    End If

End Sub
Function F_Minverse(ByVal Hairetu)
    '20210603����
    'F_Minverse(input_M)
    'F_Minverse(�z��)
    '�]���q�s���p���ċt�s����v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Hairetu, 1) <> 1 Or LBound(Hairetu, 2) <> 1 Then
        Hairetu = Application.Transpose(Application.Transpose(Hairetu))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Hairetu)
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, M2%, N% '�����グ�p(Integer�^)
    N = UBound(Hairetu, 1)
    Dim Output#()
    ReDim Output(1 To N, 1 To N)
    
    Dim detM# '�s�񎮂̒l���i�[
    detM = F_MDeterm(Hairetu) '�s�񎮂����߂�
    
    Dim Mjyokyo '�w��̗�E�s�����������z����i�[
    
    For I = 1 To N '�e��
        For J = 1 To N '�e�s
            
            'I��,J�s����������
            Mjyokyo = F_Mjyokyo(Hairetu, J, I)
            
            'I��,J�s�̗]���q�����߂ďo�͂���t�s��Ɋi�[
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_Minverse = Output
    
End Function

Function F_MDeterm(Hairetu)
    '20210603����
    'F_MDeterm(Hairetu)
    'F_MDeterm(�z��)
    '�s�񎮂��v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Hairetu, 1) <> 1 Or LBound(Hairetu, 2) <> 1 Then
        Hairetu = Application.Transpose(Application.Transpose(Hairetu))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Hairetu)
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(Hairetu, 1)
    
    Dim Hairetu2 '�|���o�����s���s��
    Hairetu2 = Hairetu
    
    For I = 1 To N '�e��
        For J = I To N '�|���o�����̍s�̒T��
            If Hairetu2(J, I) <> 0 Then
                K = J '�|���o�����̍s
                Exit For
            End If
            
            If J = N And Hairetu2(J, I) = 0 Then '�|���o�����̒l���S��0�Ȃ�s�񎮂̒l��0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I��,I�s)�ȊO�ő|���o���ƂȂ�ꍇ�͍s�����ւ�
            Hairetu2 = F_Mgyoirekae(Hairetu2, I, K)
        End If
        
        '�|���o��
        Hairetu2 = F_Mgyohakidasi(Hairetu2, I, I)
              
    Next I
    
    
    '�s�񎮂̌v�Z
    Dim Output#
    Output = 1
    
    For I = 1 To N '�e(I��,I�s)���|�����킹�Ă���
        Output = Output * Hairetu2(I, I)
    Next I
    
    '�o�́�����������������������������������������������������
    F_MDeterm = Output
    
    End Function


Function F_Mgyoirekae(Hairetu, Row1%, Row2%)
    '20210603����
    'F_Mgyoirekae(Hairetu, Row1, Row2)
    'F_Mgyoirekae(�z��,�w��s�ԍ��@,�w��s�ԍ��A)
    '�s��Hairetu�̇@�s�ƇA�s�����ւ���
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    Output = Hairetu
    M = UBound(Hairetu, 2) '�񐔎擾
    
    For I = 1 To M
        Output(Row2, I) = Hairetu(Row1, I)
        Output(Row1, I) = Hairetu(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Function F_Mgyohakidasi(Hairetu, Row%, Col%)
    '20210603����
    'F_Mgyohakidasi(Hairetu, Row, Col)
    'F_Mgyohakidasi(�z��,�w��s,�w���)
    '�s��Hairetu��Row�s�Col��̒l�Ŋe�s��|���o��
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    Output = Hairetu
    N = UBound(Output, 1) '�s���擾
    
    Dim Hakidasi '�|���o�����̍s
    Dim X# '�|���o�����̒l
    Dim Y#
    ReDim Hakidasi(1 To N)
    X = Hairetu(Row, Col)
    
    For I = 1 To N '�|���o������1�s���쐬
        Hakidasi(I) = Hairetu(Row, I)
    Next I
    
    
    For I = 1 To N '�e�s
        If I = Row Then
            '�|���o�����̍s�̏ꍇ�͂��̂܂�
            For J = 1 To N
                Output(I, J) = Hairetu(I, J)
            Next J
        
        Else
            '�|���o�����̍s�ȊO�̏ꍇ�͑|���o��
            Y = Hairetu(I, Col) '�|���o����̗�̒l
            For J = 1 To N
                Output(I, J) = Hairetu(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Function F_Mjyokyo(Hairetu, Row%, Col%)
    '20210603����
    'F_Mjyokyo(Hairetu, Row, Col)
    'F_Mjyokyo(�z��,�w��s,�w���)
    '�s��Hairetu��Row�s�ACol������������s���Ԃ�
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output '�w�肵���s�E���������̔z��
    
    N = UBound(Hairetu, 1) '�s���擾
    M = UBound(Hairetu, 2) '�񐔎擾
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2%, J2%
    
    I2 = 0 '�s���������グ������
    For I = 1 To N
        If I = Row Then
            '�Ȃɂ����Ȃ�
        Else
            I2 = I2 + 1 '�s���������グ
            
            J2 = 0 '����������グ������
            For J = 1 To M
                If J = Col Then
                    '�Ȃɂ����Ȃ�
                Else
                    J2 = J2 + 1 '����������グ
                    Output(I2, J2) = Hairetu(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function
