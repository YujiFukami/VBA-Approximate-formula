Attribute VB_Name = "ModApproximate"
'�V�[�g�֐��p�ߎ��A��Ԋ֐�
'2016/12/22�X�V
'20170622 F_SENKEIHOKAN��ǉ�
Function F_SplineXY(ByVal HairetuXY, X#)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͒lX�ɑ΂����ԒlY
    
    '�����͒l�̐�����
    'HariretuXY�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'HairetuXY��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    'X:��Ԉ��X�̒l
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(HairetuXY, 1) <> 1 Or LBound(HairetuXY, 2) <> 1 Then
        HairetuXY = Application.Transpose(Application.Transpose(HairetuXY))
    End If
    
    Dim HairetuX, HairetuY
    Dim I%, N%
    N = UBound(HairetuXY, 1)
    ReDim HairetuX(1 To N)
    ReDim HairetuY(1 To N)
    
    For I = 1 To N
        HairetuX(I) = HairetuXY(I, 1)
        HairetuY(I) = HairetuXY(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim Output_Y#
    Output_Y = F_Spline(HairetuX, HairetuY, X)
    
    '�o�́�����������������������������������������������������
    F_SplineXY = Output_Y
    
End Function

Function F_SplineXYByXList(ByVal HairetuXY, ByVal XList)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��XList�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HariretuXY�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'HairetuXY��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    'XList:��ԈʒuX���i�[���ꂽ�z��
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(HairetuXY, 1) <> 1 Or LBound(HairetuXY, 2) <> 1 Then
        HairetuXY = Application.Transpose(Application.Transpose(HairetuXY))
    End If
    
    Dim HairetuX, HairetuY
    Dim I%, N%
    N = UBound(HairetuXY, 1)
    ReDim HairetuX(1 To N)
    ReDim HairetuY(1 To N)
    
    For I = 1 To N
        HairetuX(I) = HairetuXY(I, 1)
        HairetuY(I) = HairetuXY(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim Output_YList
    Output_YList = F_SplineByXList(HairetuX, HairetuY, XList)
    
    '�o�́�����������������������������������������������������
    F_SplineXYByXList = Output_YList
    
End Function

Function F_SplineXYPara(ByVal HairetuXY, BunkatuN&)
    '�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
    'HairetuX,HairetuY���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
    '���o�͒l�̐�����
    '�p�����g���b�N�֐��`���ŕۊǂ��ꂽXList,YList���i�[���ꂽXYList
    '1��ڂ�XList,2��ڂ�YList
    
    '�����͒l�̐�����
    'HariretuXY�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'HairetuXY��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    '�p�����g���b�N�֐��̕������i�o�͂����XList,YList�̗v�f����(������+1)�j
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    Dim StartNum%
    StartNum = LBound(HairetuXY) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(HairetuXY, 1) <> 1 Or LBound(HairetuXY, 2) <> 1 Then
        HairetuXY = Application.Transpose(Application.Transpose(HairetuXY))
    End If
    
    Dim HairetuX, HairetuY
    Dim I%, N%
    N = UBound(HairetuXY, 1)
    ReDim HairetuX(StartNum To StartNum - 1 + N)
    ReDim HairetuY(StartNum To StartNum - 1 + N)
    
    For I = 1 To N
        HairetuX(I + StartNum - 1) = HairetuXY(I, 1)
        HairetuY(I + StartNum - 1) = HairetuXY(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim Dummy
    Dim Output_XList, Output_YList
    Dummy = F_SplinePara(HairetuX, HairetuY, BunkatuN)
    Output_XList = Dummy(1)
    Output_YList = Dummy(2)
    
    Dim OutputXYList
    ReDim OutputXYList(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 2)
    
    For I = 1 To BunkatuN + 1
        OutputXYList(StartNum + I - 1, 1) = Output_XList(StartNum + I - 1)
        OutputXYList(StartNum + I - 1, 2) = Output_YList(StartNum + I - 1)
    Next I
    
    '�o�́�����������������������������������������������������
    F_SplineXYPara = OutputXYList
    
End Function

Function F_Spline#(ByVal HairetuX, ByVal HairetuY, X#)
        
    '20171124�C��
    '20180309����
    
    '�X�v���C����Ԍv�Z���s��
    '
    '<�o�͒l�̐���>
    '���͒lX�ɑ΂����ԒlY
    '
    '<���͒l�̐���>
    'HairetuX�F��Ԃ̑ΏۂƂ���z��X
    'HairetuY�F��Ԃ̑ΏۂƂ���z��Y
    'X�F��Ԉʒu��X�̒l
    
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(HairetuX, 1) <> 1 Then
        HairetuX = Application.Transpose(Application.Transpose(HairetuX))
    End If
    If LBound(HairetuY, 1) <> 1 Then
        HairetuY = Application.Transpose(Application.Transpose(HairetuY))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(HairetuX, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(HairetuY, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        HairetuX = Application.Transpose(HairetuX)
    End If
    If JigenCheck2 > 0 Then
        HairetuY = Application.Transpose(HairetuY)
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I%, N%, K%, A, B, C, D
    Dim Output_Y# '�o�͒lY
    Dim SotoNaraTrue As Boolean
    SotoNaraTrue = False
    N = UBound(HairetuX, 1)
       
    '�X�v���C���v�Z�p�̊e�W�����v�Z����B�Q�Ɠn����A,B,C,D�Ɋi�[
    Dim Dummy
    Dummy = SplineKeisu(HairetuX, HairetuY)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
        
    For I = 1 To N - 1
        If HairetuX(I) < HairetuX(I + 1) Then 'X���P�������̏ꍇ
            If I = 1 And HairetuX(1) > X Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                Output_Y = HairetuY(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And HairetuX(I + 1) <= X Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                Output_Y = HairetuY(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf HairetuX(I) <= X And HairetuX(I + 1) > X Then '�͈͓�
                K = I: Exit For
            
            End If
        Else 'X���P�������̏ꍇ
        
            If I = 1 And HairetuX(1) < X Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                Output_Y = HairetuY(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And HairetuX(I + 1) >= X Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                Output_Y = HairetuY(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf HairetuX(I + 1) < X And HairetuX(I) >= X Then '�͈͓�
                K = I: Exit For
            
            End If
        
        End If
    Next I
        
    If SotoNaraTrue = False Then
        Output_Y = A(K) + B(K) * (X - HairetuX(K)) + C(K) * (X - HairetuX(K)) ^ 2 + D(K) * (X - HairetuX(K)) ^ 3
    End If
    
    '�o�́�����������������������������������������������������
    F_Spline = Output_Y

End Function

Function F_SplinePara(ByVal HairetuX, ByVal HairetuY, BunkatuN&)
    '�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
    'HairetuX,HairetuY���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
    '���o�͒l�̐�����
    '�p�����g���b�N�֐��`���ŕۊǂ��ꂽXList,YList
    
    '�����͒l�̐�����
    'HariretuX�F��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
    'HariretuY�F��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
    '�p�����g���b�N�֐��̕������i�o�͂����XList,YList�̗v�f����(������+1)�j
    
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    Dim StartNum%
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    StartNum = LBound(HairetuX, 1) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(HairetuX, 1) <> 1 Then
        HairetuX = Application.Transpose(Application.Transpose(HairetuX))
    End If
    If LBound(HairetuY, 1) <> 1 Then
        HairetuY = Application.Transpose(Application.Transpose(HairetuY))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(HairetuX, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(HairetuY, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        HairetuX = Application.Transpose(HairetuX)
    End If
    If JigenCheck2 > 0 Then
        HairetuY = Application.Transpose(HairetuY)
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(HairetuX, 1)
    Dim HairetuT#(), TList#()
    
    'X,Y�̕�Ԃ̊�ƂȂ�z����쐬
    ReDim HairetuT(1 To N)
    For I = 1 To N
        '0�`1�𓙊Ԋu
        HairetuT(I) = (I - 1) / (N - 1)
    Next I
    
    '�o�͕�Ԉʒu�̊�ʒu
    If JigenCheck1 > 0 Then '�o�͒l�̌`�����͒l�ɍ��킹�邽�߂̏���
        ReDim TList(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            TList(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim TList(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            TList(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim Output_XList, Output_YList
    Output_XList = F_SplineByXList(HairetuT, HairetuX, TList)
    Output_YList = F_SplineByXList(HairetuT, HairetuY, TList)
    
    '�o��
    Dim Output(1 To 2)
    Output(1) = Output_XList
    Output(2) = Output_YList
    
    F_SplinePara = Output
    
End Function

Function F_SplineByXList(ByVal HairetuX, ByVal HairetuY, ByVal XList)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��XList�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HariretuX�F��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
    'HariretuY�F��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
    'XList:��ԈʒuX���i�[���ꂽ�z��

    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    Dim StartNum%
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(HairetuX, 1) <> 1 Then
        HairetuX = Application.Transpose(Application.Transpose(HairetuX))
    End If
    If LBound(HairetuY, 1) <> 1 Then
        HairetuY = Application.Transpose(Application.Transpose(HairetuY))
    End If
    StartNum = LBound(XList, 1) 'XList�̊J�n�v�f�ԍ�������Ă����i�o�͒l�����킹�邽�߁j
    If LBound(XList, 1) <> 1 Then
        XList = Application.Transpose(Application.Transpose(XList))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(HairetuX, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(HairetuY, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck3 = UBound(XList, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        HairetuX = Application.Transpose(HairetuX)
    End If
    If JigenCheck2 > 0 Then
        HairetuY = Application.Transpose(HairetuY)
    End If
    If JigenCheck3 > 0 Then
        XList = Application.Transpose(XList)
    End If

    '�v�Z����������������������������������������������������������
    Dim N%, K%, A, B, C, D
    
    '�X�v���C���v�Z�p�̊e�W�����v�Z����B�Q�Ɠn����A,B,C,D�Ɋi�[
    Dim Dummy
    Dummy = SplineKeisu(HairetuX, HairetuY)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(HairetuX, 1) '��ԑΏۂ̗v�f��
    
    Dim Output_YList#() '�o�͂���YList�̊i�[
    Dim NX%
    NX = UBound(XList, 1) '��Ԉʒu�̌�
    ReDim Output_YList(1 To NX)
    Dim TmpX#, TmpY#
    
    For J = 1 To NX
        TmpX = XList(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If HairetuX(I) < HairetuX(I + 1) Then 'X���P�������̏ꍇ
                If I = 1 And HairetuX(1) > TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = HairetuY(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And HairetuX(I + 1) <= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = HairetuY(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf HairetuX(I) <= TmpX And HairetuX(I + 1) > TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            Else 'X���P�������̏ꍇ
            
                If I = 1 And HairetuX(1) < TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = HairetuY(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And HairetuX(I + 1) >= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = HairetuY(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf HairetuX(I + 1) < TmpX And HairetuX(I) >= TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - HairetuX(K)) + C(K) * (TmpX - HairetuX(K)) ^ 2 + D(K) * (TmpX - HairetuX(K)) ^ 3
        End If
        
        Output_YList(J) = TmpY
        
    Next J
    
    '�o�́�����������������������������������������������������
    Dim Output
    
    '�o�͂���z�����͂����z��XList�̌`��ɍ��킹��
    If JigenCheck3 = 1 Then '���͂�XList���񎟌��z��
        ReDim Output(StartNum To StartNum + NX - 1, 1 To 1)
        For I = 1 To NX
            Output(StartNum + I - 1, 1) = Output_YList(I)
        Next I
    Else
        If StartNum = 1 Then
            Output = Output_YList
        Else
            ReDim Output(StartNum To StartNum + NX - 1)
            For I = 1 To NX
                Output(StartNum + I - 1) = Output_YList(I)
            Next I
        End If
    End If
    
    F_SplineByXList = Output
    
End Function

Function SplineKeisu(ByVal X, ByVal Y)

    '�Q�l�Fhttp://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim A, B, C, D
    N = UBound(X, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h#()
    Dim Hairetu_L#() '���ӂ̔z�� �v�f��(1 to N,1 to N)
    Dim Hairetu_R#() '�E�ӂ̔z�� �v�f��(1 to N,1 to 1)
    Dim Hairetu_Lm#() '���ӂ̔z��̋t�s�� �v�f��(1 to N,1 to N)
    
    ReDim h(1 To N - 1)
    ReDim Hairetu_L(1 To N, 1 To N)
    ReDim Hairetu_R(1 To N, 1 To 1)
    
    'hi = xi+1 - x
    For I = 1 To N - 1
        h(I) = X(I + 1) - X(I)
    Next I
    
    'di = yi
    For I = 1 To N
        A(I) = Y(I)
    Next I
    
    '�E�ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Or I = N Then
            Hairetu_R(I, 1) = 0
        Else
            Hairetu_R(I, 1) = 3 * (Y(I + 1) - Y(I)) / h(I) - 3 * (Y(I) - Y(I - 1)) / h(I - 1)
        End If
    Next I
    
    '���ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Then
            Hairetu_L(I, 1) = 1
        ElseIf I = N Then
            Hairetu_L(N, N) = 1
        Else
            Hairetu_L(I - 1, I) = h(I - 1)
            Hairetu_L(I, I) = 2 * (h(I) + h(I - 1))
            Hairetu_L(I + 1, I) = h(I)
        End If
    Next I
    
    '���ӂ̔z��̋t�s��
    Hairetu_Lm = F_Minverse(Hairetu_L)
    
    'C�̔z������߂�
    C = F_MMult(Hairetu_Lm, Hairetu_R)
    C = Application.Transpose(C)
    
    'B�̔z������߂�
    For I = 1 To N - 1
        B(I) = (A(I + 1) - A(I)) / h(I) - h(I) * (C(I + 1) + 2 * C(I)) / 3
    Next I
    
    'D�̔z������߂�
    For I = 1 To N - 1
        D(I) = (C(I + 1) - C(I)) / (3 * h(I))
    Next I
    
    '�o��
    Dim Output(1 To 4)
    Output(1) = A
    Output(2) = B
    Output(3) = C
    Output(4) = D
    
    SplineKeisu = Output

End Function
