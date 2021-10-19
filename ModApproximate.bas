Attribute VB_Name = "ModApproximate"
'�V�[�g�֐��p�ߎ��A��Ԋ֐�
Private Sub TestSplineXY()
'SplineXY�̎��s�e�X�g

    Dim ArrayXY2D
    Dim InputX   As Double
    Dim OutputY  As Double
    ArrayXY2D = Application.Transpose(Application.Transpose( _
                Array(Array(0, 5.93255769237665), _
                Array(1, 9.99308268536779), _
                Array(2, 5.5044328013839), _
                Array(3, 5.60877013928983), _
                Array(4, 1.51682123665907), _
                Array(5, 8.18738634627902), _
                Array(6, 4.42233332813268)) _
                ))
    
    InputX = 3.5
    OutputY = SplineXY(ArrayXY2D, InputX)
    
    Debug.Print "OutputY = " & OutputY
    
End Sub

Private Sub TestSpline()
'Spline�̎��s�e�X�g

    Dim ArrayX1D
    Dim ArrayY1D
    Dim InputX As Double
    ArrayX1D = Application.Transpose(Application.Transpose( _
                Array(0, 1, 2, 3, 4, 5, 6) _
                ))
    
    ArrayY1D = Application.Transpose(Application.Transpose( _
                Array(5.93255769237665, 9.99308268536779, 5.5044328013839, 5.60877013928983, 1.51682123665907, 8.18738634627902, 4.42233332813268) _
                ))
                
    InputX = 3.5
    
    Dim OutputY As Double
    OutputY = Spline(ArrayX1D, ArrayY1D, InputX)

    Debug.Print "OutputY = " & OutputY
    
End Sub

Private Sub TestSplineXYByArrayX1D()
'SplineXYByArrayX1D�̎��s�e�X�g

    Dim ArrayXY2D
    Dim InputArrayX1D
    ArrayXY2D = Application.Transpose(Application.Transpose( _
                Array(Array(0, 5.93255769237665), _
                Array(1, 9.99308268536779), _
                Array(2, 5.5044328013839), _
                Array(3, 5.60877013928983), _
                Array(4, 1.51682123665907), _
                Array(5, 8.18738634627902), _
                Array(6, 4.42233332813268)) _
                ))
    
    InputArrayX1D = Application.Transpose(Application.Transpose( _
                    Array(0.704709737423495, 1.15605119826871, 1.68490822086298, 2.13925473863431, 2.58350091448881, 3.13230954582088, 3.27625171436593, 3.96995547976061, 4.5878879819556, 5.29470346416526) _
                    ))
    
    
    Dim OutputArrayY1D
    OutputArrayY1D = SplineXYByArrayX1D(ArrayXY2D, InputArrayX1D)
    
    Call DPH(OutputArrayY1D)

End Sub

Private Sub TestSplineByArrayX1D()
'SplineByArrayX1D�̎��s�e�X�g

    Dim ArrayX1D
    Dim ArrayY1D
    Dim InputArrayX1D
    ArrayX1D = Application.Transpose(Application.Transpose( _
                Array(0, 1, 2, 3, 4, 5, 6) _
                ))
    
    ArrayY1D = Application.Transpose(Application.Transpose( _
                Array(5.93255769237665, 9.99308268536779, 5.5044328013839, 5.60877013928983, 1.51682123665907, 8.18738634627902, 4.42233332813268) _
                ))
    
    InputArrayX1D = Application.Transpose(Application.Transpose( _
                    Array(0.704709737423495, 1.15605119826871, 1.68490822086298, 2.13925473863431, 2.58350091448881, 3.13230954582088, 3.27625171436593, 3.96995547976061, 4.5878879819556, 5.29470346416526) _
                    ))
    
    Dim OutputArrayY1D
    OutputArrayY1D = SplineByArrayX1D(ArrayX1D, ArrayY1D, InputArrayX1D)
    
    Call DPH(OutputArrayY1D)

End Sub

Private Sub TestSplineXYPara()
'SplineXYPara�̎��s�e�X�g

    Dim ArrayXY2D
    Dim BunkatuN As Long
    ArrayXY2D = Application.Transpose(Application.Transpose( _
                Array(Array(0, 5.93255769237665), _
                Array(1, 9.99308268536779), _
                Array(2, 5.5044328013839), _
                Array(3, 5.60877013928983), _
                Array(4, 1.51682123665907), _
                Array(5, 8.18738634627902), _
                Array(6, 4.42233332813268)) _
                ))
                
    BunkatuN = 10
    
    Dim OutputArrayXY2D
    
    OutputArrayXY2D = SplineXYPara(ArrayXY2D, BunkatuN)
    
    Call DPH(OutputArrayXY2D)
    
End Sub

Private Sub TestSplinePara()
'SplinePara�̎��s�e�X�g

    Dim ArrayX1D
    Dim ArrayY1D
    Dim BunkatuN As Long
    ArrayX1D = Application.Transpose(Application.Transpose( _
                Array(0, 1, 2, 3, 4, 5, 6) _
                ))
    
    ArrayY1D = Application.Transpose(Application.Transpose( _
                Array(5.93255769237665, 9.99308268536779, 5.5044328013839, 5.60877013928983, 1.51682123665907, 8.18738634627902, 4.42233332813268) _
                ))
    
    BunkatuN = 10
    
    Dim OutputArrayY1D
    OutputArrayY1D = SplinePara(ArrayX1D, ArrayY1D, BunkatuN)
    
    Call DPH(OutputArrayY1D)
    
End Sub

Function SplineXY(ByVal ArrayXY2D, InputX As Double)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͒lX�ɑ΂����ԒlY
    
    '�����͒l�̐�����
    'ArrayXY2D�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'ArrayXY2D��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    'X:��Ԉʒu��X�̒l
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean: RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D
    Dim ArrayY1D
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(1 To N)
    ReDim ArrayY1D(1 To N)
    
    For I = 1 To N
        ArrayX1D(I) = ArrayXY2D(I, 1)
        ArrayY1D(I) = ArrayXY2D(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim OutputY As Double
    OutputY = Spline(ArrayX1D, ArrayY1D, InputX)
    
    '�o�́�����������������������������������������������������
    If RangeNaraTrue Then
        '���[�N�V�[�g�֐��̏ꍇ
        SplineXY = Application.Transpose(OutputY)
    Else
        'VBA��ł̏����̏ꍇ
        SplineXY = OutputY
    End If
    
End Function

Function SplineXYByArrayX1D(ByVal ArrayXY2D, ByVal InputArrayX1D)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HariretuXY�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'ArrayXY2D��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    'InputArrayX1D:��ԈʒuX���i�[���ꂽ�z��
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean: RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If

    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D
    Dim ArrayY1D
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(1 To N)
    ReDim ArrayY1D(1 To N)
    
    For I = 1 To N
        ArrayX1D(I) = ArrayXY2D(I, 1)
        ArrayY1D(I) = ArrayXY2D(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim OutputArrayY1D
    OutputArrayY1D = SplineByArrayX1D(ArrayX1D, ArrayY1D, InputArrayX1D)
    
    '�o�́�����������������������������������������������������
    If RangeNaraTrue = True Then
        '���[�N�V�[�g�֐��̏ꍇ
        SplineXYByArrayX1D = Application.Transpose(OutputArrayY1D)
    Else
        'VBA��ł̏����̏ꍇ
        SplineXYByArrayX1D = OutputArrayY1D
    End If
    
End Function

Function SplineXYPara(ByVal ArrayXY2D, BunkatuN As Long)
    '�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
    'ArrayX,ArrayY���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
    '���o�͒l�̐�����
    '�p�����g���b�N�֐��`���ŕ�Ԃ��ꂽXList,YList���i�[���ꂽXYList
    '1��ڂ�XList,2��ڂ�YList
    
    '�����͒l�̐�����
    'ArrayXY2D�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'ArrayXY2D��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    '�p�����g���b�N�֐��̕������i�o�͂����XList,YList�̗v�f����(������+1)�j
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
    End If
        
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    Dim StartNum As Integer
    StartNum = LBound(ArrayXY2D) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D
    Dim ArrayY1D
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(StartNum To StartNum - 1 + N)
    ReDim ArrayY1D(StartNum To StartNum - 1 + N)
    
    For I = 1 To N
        ArrayX1D(I + StartNum - 1) = ArrayXY2D(I, 1)
        ArrayY1D(I + StartNum - 1) = ArrayXY2D(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim Dummy
    Dim OutputArrayX1D
    Dim OutputArrayY1D
    Dummy = SplinePara(ArrayX1D, ArrayY1D, BunkatuN)
    OutputArrayX1D = Dummy(1)
    OutputArrayY1D = Dummy(2)
    
    Dim OutputArrayXY2D
    ReDim OutputArrayXY2D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 2)
    
    For I = 1 To BunkatuN + 1
        OutputArrayXY2D(StartNum + I - 1, 1) = OutputArrayX1D(StartNum + I - 1)
        OutputArrayXY2D(StartNum + I - 1, 2) = OutputArrayY1D(StartNum + I - 1)
    Next I
    
    '�o�́�����������������������������������������������������
    SplineXYPara = OutputArrayXY2D
    
End Function

Function SplineXYParaFast(ByVal ArrayXY2D, BunkatuN As Long, PointCount As Long)
'�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
'�������Čv�Z������������
'ArrayX,ArrayY���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
    
'����
'ArrayXY2D �E�E�E��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
'ArrayXY2D��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
'BunkatuN  �E�E�E�p�����g���b�N�֐��̕������i�o�͂����XList,YList�̗v�f����(������+1)�j
'PointCount�E�E�E��������ۂ̓_��
    
'�Ԃ�l
'�p�����g���b�N�֐��`���ŕ�Ԃ��ꂽXList,YList���i�[���ꂽXYList
'1��ڂ�XList,2��ڂ�YList
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
    End If
        
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    Dim StartNum As Integer
    StartNum = LBound(ArrayXY2D) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D
    Dim ArrayY1D
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(StartNum To StartNum - 1 + N)
    ReDim ArrayY1D(StartNum To StartNum - 1 + N)
    
    For I = 1 To N
        ArrayX1D(I + StartNum - 1) = ArrayXY2D(I, 1)
        ArrayY1D(I + StartNum - 1) = ArrayXY2D(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim Dummy
    Dim OutputArrayX1D
    Dim OutputArrayY1D
    Dummy = SplineParaFast(ArrayX1D, ArrayY1D, BunkatuN, PointCount)
    OutputArrayX1D = Dummy(1)
    OutputArrayY1D = Dummy(2)
    
    Dim OutputArrayXY2D
    ReDim OutputArrayXY2D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 2)
    
    For I = 1 To BunkatuN + 1
        OutputArrayXY2D(StartNum + I - 1, 1) = OutputArrayX1D(StartNum + I - 1)
        OutputArrayXY2D(StartNum + I - 1, 2) = OutputArrayY1D(StartNum + I - 1)
    Next I
    
    '�o�́�����������������������������������������������������
    SplineXYParaFast = OutputArrayXY2D
    
End Function


Function Spline(ByVal ArrayX1D, ByVal ArrayY1D, InputX As Double)
        
    '20171124�C��
    '20180309����
    
    '�X�v���C����Ԍv�Z���s��
    
    '<�o�͒l�̐���>
    '���͒lX�ɑ΂����ԒlY
    
    '<���͒l�̐���>
    'ArrayX1D�F��Ԃ̑ΏۂƂ���z��X
    'ArrayY1D�F��Ԃ̑ΏۂƂ���z��Y
    'InputX  �F��Ԉʒu��X�̒l
    
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I As Integer
    Dim N As Integer
    Dim K As Integer
    Dim A
    Dim B
    Dim C
    Dim D
    Dim OutputY As Double '�o�͒lY
    Dim SotoNaraTrue As Boolean
    SotoNaraTrue = False
    N = UBound(ArrayX1D, 1)
       
    '�X�v���C���v�Z�p�̊e�W�����v�Z����B�Q�Ɠn����A,B,C,D�Ɋi�[
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
        
    For I = 1 To N - 1
        If ArrayX1D(I) < ArrayX1D(I + 1) Then 'X���P�������̏ꍇ
            If I = 1 And ArrayX1D(1) > InputX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                OutputY = ArrayY1D(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And ArrayX1D(I + 1) <= InputX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                OutputY = ArrayY1D(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf ArrayX1D(I) <= InputX And ArrayX1D(I + 1) > InputX Then '�͈͓�
                K = I: Exit For
            
            End If
        Else 'X���P�������̏ꍇ
        
            If I = 1 And ArrayX1D(1) < InputX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                OutputY = ArrayY1D(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And ArrayX1D(I + 1) >= InputX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                OutputY = ArrayY1D(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf ArrayX1D(I + 1) < InputX And ArrayX1D(I) >= InputX Then '�͈͓�
                K = I: Exit For
            
            End If
        
        End If
    Next I
        
    If SotoNaraTrue = False Then
        OutputY = A(K) + B(K) * (InputX - ArrayX1D(K)) + C(K) * (InputX - ArrayX1D(K)) ^ 2 + D(K) * (InputX - ArrayX1D(K)) ^ 3
    End If
    
    '�o�́�����������������������������������������������������
    Spline = OutputY

End Function

Function SplinePara(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN As Long)
    '�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
    'ArrayX1D,ArrayY1D���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
    '���o�͒l�̐�����
    '�p�����g���b�N�֐��`���ŕ�Ԃ��ꂽXList,YList
    
    '�����͒l�̐�����
    'ArrayX1D�F��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
    'ArrayY1D�F��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
    '�p�����g���b�N�֐��̕������i�o�͂����OutputArrayX1D,OutputArrayY1D�̗v�f����(������+1)�j
    
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    Dim StartNum As Integer
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    StartNum = LBound(ArrayX1D, 1) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D()     As Double
    Dim ArrayParaT1D() As Double
    
    'X,Y�̕�Ԃ̊�ƂȂ�z����쐬
    ReDim ArrayT1D(1 To N)
    For I = 1 To N
        '0�`1�𓙊Ԋu
        ArrayT1D(I) = (I - 1) / (N - 1)
    Next I
    
    '�o�͕�Ԉʒu�̊�ʒu
    If JigenCheck1 > 0 Then '�o�͒l�̌`�����͒l�ɍ��킹�邽�߂̏���
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            ArrayParaT1D(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            ArrayParaT1D(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim OutputArrayX1D
    Dim OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1D(ArrayT1D, ArrayX1D, ArrayParaT1D)
    OutputArrayY1D = SplineByArrayX1D(ArrayT1D, ArrayY1D, ArrayParaT1D)
    
    '�o��
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplinePara = Output
    
End Function

Function SplineParaFast(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN As Long, PointCount As Long)
'�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
'�������Čv�Z������������
'ArrayX1D,ArrayY1D���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
'20211009

'����
'ArrayX1D  �E�E�E��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
'ArrayY1D  �E�E�E��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
'BunkatuN  �E�E�E�p�����g���b�N�֐��̕������i�o�͂����OutputArrayX1D,OutputArrayY1D�̗v�f����(������+1)�j
'PointCount�E�E�E��������ۂ̓_��

'�Ԃ�l
'�p�����g���b�N�֐��`���ŕ�Ԃ��ꂽXList,YList
    
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    Dim StartNum As Integer
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    StartNum = LBound(ArrayX1D, 1) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D() As Double, ArrayParaT1D() As Double
    
    'X,Y�̕�Ԃ̊�ƂȂ�z����쐬
    ReDim ArrayT1D(1 To N)
    For I = 1 To N
        '0�`1�𓙊Ԋu
        ArrayT1D(I) = (I - 1) / (N - 1)
    Next I
    
    '�o�͕�Ԉʒu�̊�ʒu
    If JigenCheck1 > 0 Then '�o�͒l�̌`�����͒l�ɍ��킹�邽�߂̏���
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            ArrayParaT1D(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            ArrayParaT1D(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim OutputArrayX1D
    Dim OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1DFast(ArrayT1D, ArrayX1D, ArrayParaT1D, PointCount)
    OutputArrayY1D = SplineByArrayX1DFast(ArrayT1D, ArrayY1D, ArrayParaT1D, PointCount)
    
    '�o��
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplineParaFast = Output
    
End Function


Private Function �X�v���C����ԍ������p�ɕ�������(ByVal ArrayX1D, ByVal ArrayY1D, ByVal CalPoint1D, PointCount As Long)
'�X�v���C����ԍ������p�ɕ�������
'20211009

'����
'ArrayX1D  �E�E�E��Ԍ���X���W���X�g
'ArrayY1D  �E�E�E��Ԍ���Y���W���X�g
'CalPoint1D�E�E�E��Ԉʒu��X���W���X�g
'PointCount�E�E�E������̈�̕����̓_��

    Dim I  As Long
    Dim J  As Long
    Dim II As Long
    Dim JJ As Long
    Dim N  As Long
    Dim M  As Long
    Dim K  As Long
    N = UBound(ArrayX1D, 1)
    Dim PointN As Long
    PointN = UBound(CalPoint1D, 1)
    
    Dim Output '�o�͒l�i�[�ϐ�
    ReDim Output(1 To N, 1 To 3) '1:��Ԍ�X���W���X�g,2:��Ԍ�Y���W���X�g,3:��ԈʒuX���W���X�g
    'N�͂Ƃ肠�����̍ő�ŁA��Ŕz����k������
    
    Dim TmpXList
    Dim TmpYList
    Dim TmpPointList
    Dim TmpInterXList
    Dim StartNum      As Long '���������Ԍ����W�̊J�n�ʒu
    Dim EndNum        As Long '���������Ԍ����W�̏I���ʒu
    Dim InterStartNum As Long '�������ꂽ��Ԍ����W�Ŏ��ۂ̕�Ԕ͈͂̊J�n�ʒu
    Dim InterEndNum   As Long '�������ꂽ��Ԍ����W�Ŏ��ۂ̕�Ԕ͈͂̏I���ʒu
    
    K = 0
    Do
        K = K + 1
        StartNum = (K - 1) * PointCount - 2
        EndNum = StartNum + PointCount + 2
        If StartNum <= 1 Then
            InterStartNum = 1
            StartNum = 1
        Else
            InterStartNum = StartNum + 1
        End If
        
        If EndNum >= N Then
            InterEndNum = N
            EndNum = N
        Else
            InterEndNum = EndNum - 1
        End If
        
        TmpXList = ExtractArray1D(ArrayX1D, StartNum, EndNum)
        TmpYList = ExtractArray1D(ArrayY1D, StartNum, EndNum)
        TmpInterXList = ExtractArray1D(ArrayX1D, InterStartNum, InterEndNum)
        TmpPointList = ExtractByRangeArray1D(CalPoint1D, TmpInterXList)
        
        Output(K, 1) = TmpXList
        Output(K, 2) = TmpYList
        Output(K, 3) = TmpPointList
        
        If EndNum = N Then
            Exit Do
        End If
    Loop
    
    '�o�͂���i�[�z��͈̔͒���
    Output = ExtractArray(Output, 1, 1, K, 3)
    
    '����������Ԉʒu�ŏd��������̂�����
    N = UBound(Output, 1)
    Dim TmpList1
    Dim TmpList2
    For I = 2 To N
        TmpList1 = Output(I - 1, 3)
        TmpList2 = Output(I, 3)
        If IsEmpty(TmpList1) = False And IsEmpty(TmpList2) = False Then
            If TmpList1(UBound(TmpList1, 1)) = TmpList2(1) Then '�Ō�̗v�f�ƍŏ��̗v�f���r����
                If UBound(TmpList2, 1) = 1 Then
                    TmpList2 = Empty
                Else
                    TmpList2 = ExtractArray1D(TmpList2, 2, UBound(TmpList2, 1))
                End If
                Output(I, 3) = TmpList2
            End If
        End If
    Next
    
    �X�v���C����ԍ������p�ɕ������� = Output
    
End Function

Function ExtractByRangeArray1D(InputArray1D, RangeArray1D)
'�ꎟ���z��̎w��͈͂𒊏o����B
'�w��͈͂�RangeArray1D�Ŏw�肷��B
'20211009

'����
'InputArray1D�E�E�E���o���̈ꎟ���z��
'RangeArray1D�E�E�E���o����͈͂��w�肷��ꎟ���z��

'��
'InputArray1D = (1,2,3,4,5,6,7,8,9,10)
'RangeArray1D = (3,4,7)
'�o�� = (3,4,5,6,7)

    '�����`�F�b�N
    Call CheckArray1D(InputArray1D, "InputArray1D")
    Call CheckArray1DStart1(InputArray1D, "InputArray1D")
    Call CheckArray1D(RangeArray1D, "RangeArray1D")
    Call CheckArray1DStart1(RangeArray1D, "RangeArray1D")
    
    Dim I  As Long
    Dim J  As Long
    Dim II As Long
    Dim JJ As Long
    Dim N  As Long
    Dim M  As Long
    Dim K  As Long
    
    '�w��͈͂̍ŏ��A�ő���擾
    Dim MinNum As Double
    Dim MaxNum As Double
    MinNum = WorksheetFunction.Min(RangeArray1D)
    MaxNum = WorksheetFunction.Max(RangeArray1D)
    
    '���o�͈͂̊J�n�ʒu�A�I���ʒu���v�Z
    Dim StartNum As Long
    Dim EndNum   As Long
    StartNum = 0
    EndNum = 0
    N = UBound(InputArray1D, 1)
    For I = 1 To N
        If InputArray1D(I) >= MinNum Then
            StartNum = I
            Exit For
        End If
    Next
    
    If StartNum = 0 Then
        '���o�͈͂Ȃ���Empty��Ԃ�
        Exit Function
    End If
    
    For I = StartNum To N
        If InputArray1D(I) > MaxNum Then
            EndNum = I - 1
            Exit For
        End If
    Next
    
    If EndNum = 0 Then
        '�I���ʒu��������Ȃ��ꍇ�͏I���܂őS���܂�
        EndNum = N
    End If
    
    '�͈͒��o
    Dim Output '�o�͒l�i�[�ϐ�
    Output = ExtractArray1D(InputArray1D, StartNum, EndNum)
    
    '�o��
    ExtractByRangeArray1D = Output
    
End Function

Function SplineByArrayX1DFast(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D, PointCount As Long)
 '�X�v���C����Ԍv�Z���s��
 '�������Čv�Z���邱�Ƃō���������

'����
'HairetuX     �E�E�E��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
'HairetuY     �E�E�E��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
'InputArrayX1D�E�E�E��ԈʒuX���i�[���ꂽ�z��
'PointCount   �E�E�E��������ۂ̓_��

'�Ԃ�l
'���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��
        
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
        RangeNaraTrue = True
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If
    
    Dim StartNum As Integer
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1D�̊J�n�v�f�ԍ�������Ă����i�o�͒l�����킹�邽�߁j
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    Dim JigenCheck3 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck3 = UBound(InputArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '�v�Z����������������������������������������������������������
    Dim SplitArrayList
    SplitArrayList = �X�v���C����ԍ������p�ɕ�������(ArrayX1D, ArrayY1D, InputArrayX1D, PointCount)
        
    Dim TmpXList
    Dim TmpYList
    Dim TmpPointList
    Dim Output '�o�͒l�i�[�ϐ�
    Dim TmpSplineList
    Dim I  As Long
    Dim J  As Long
    Dim II As Long
    Dim JJ As Long
    Dim N  As Long
    Dim M  As Long
    Dim K  As Long
    N = UBound(SplitArrayList, 1)
    K = 0
    For I = 1 To N
        TmpXList = SplitArrayList(I, 1)
        TmpYList = SplitArrayList(I, 2)
        TmpPointList = SplitArrayList(I, 3)
        If IsEmpty(TmpPointList) = False Then
            TmpSplineList = SplineByArrayX1D(TmpXList, TmpYList, TmpPointList)
            K = K + 1
            If K = 1 Then
                Output = TmpSplineList
            Else
                Output = UnionArray1D(Output, TmpSplineList)
            End If
        End If
    Next
    
    SplineByArrayX1DFast = Output
    
End Function

Function SplineByArrayX1D(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HairetuX�F��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
    'HairetuY�F��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
    'InputArrayX1D:��ԈʒuX���i�[���ꂽ�z��

    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
        RangeNaraTrue = True
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If
    
    Dim StartNum As Integer
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1D�̊J�n�v�f�ԍ�������Ă����i�o�͒l�����킹�邽�߁j
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    Dim JigenCheck3 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck3 = UBound(InputArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '�v�Z����������������������������������������������������������
    Dim A, B, C, D
    Dim I As Long, J As Long, K As Long, M As Long, N As Long '�����グ�p(Long�^)
    
    '�X�v���C���v�Z�p�̊e�W�����v�Z����B�Q�Ɠn����A,B,C,D�Ɋi�[
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(ArrayX1D, 1) '��ԑΏۂ̗v�f��
    
    Dim OutputArrayY1D() As Double '�o�͂���Y�̊i�[
    Dim NX As Integer
    NX = UBound(InputArrayX1D, 1) '��Ԉʒu�̌�
    ReDim OutputArrayY1D(1 To NX)
    Dim TmpX As Double, TmpY As Double
    
    For J = 1 To NX
        TmpX = InputArrayX1D(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If ArrayX1D(I) < ArrayX1D(I + 1) Then 'X���P�������̏ꍇ
                If I = 1 And ArrayX1D(1) > TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) <= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I) <= TmpX And ArrayX1D(I + 1) > TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            Else 'X���P�������̏ꍇ
            
                If I = 1 And ArrayX1D(1) < TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) >= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I + 1) < TmpX And ArrayX1D(I) >= TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - ArrayX1D(K)) + C(K) * (TmpX - ArrayX1D(K)) ^ 2 + D(K) * (TmpX - ArrayX1D(K)) ^ 3
        End If
        
        OutputArrayY1D(J) = TmpY
        
    Next J
    
    '�o�́�����������������������������������������������������
    Dim Output
    
    '�o�͂���z�����͂����z��InputArrayX1D�̌`��ɍ��킹��
    If JigenCheck3 = 1 Then '���͂�InputArrayX1D���񎟌��z��
        ReDim Output(StartNum To StartNum + NX - 1, 1 To 1)
        For I = 1 To NX
            Output(StartNum + I - 1, 1) = OutputArrayY1D(I)
        Next I
    Else
        If StartNum = 1 Then
            Output = OutputArrayY1D
        Else
            ReDim Output(StartNum To StartNum + NX - 1)
            For I = 1 To NX
                Output(StartNum + I - 1) = OutputArrayY1D(I)
            Next I
        End If
    End If
    
    If RangeNaraTrue Then
        '���[�N�V�[�g�֐��̏ꍇ
        SplineByArrayX1D = Application.Transpose(Output)
    Else
        'VBA��ł̏����̏ꍇ
        SplineByArrayX1D = Output
    End If
    
End Function

Function SplineKeisu(ByVal ArrayX1D, ByVal ArrayY1D)

    '�Q�l�Fhttp://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I As Integer
    Dim N As Integer
    Dim A
    Dim B
    Dim C
    Dim D
    N = UBound(ArrayX1D, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h()         As Double
    Dim ArrayL2D()  As Double '���ӂ̔z�� �v�f��(1 to N,1 to N)
    Dim ArrayR1D()  As Double '�E�ӂ̔z�� �v�f��(1 to N,1 to 1)
    Dim ArrayLm2D() As Double '���ӂ̔z��̋t�s�� �v�f��(1 to N,1 to N)
    
    ReDim h(1 To N - 1)
    ReDim ArrayL2D(1 To N, 1 To N)
    ReDim ArrayR1D(1 To N, 1 To 1)
    
    'hi = xi+1 - x
    For I = 1 To N - 1
        h(I) = ArrayX1D(I + 1) - ArrayX1D(I)
    Next I
    
    'di = yi
    For I = 1 To N
        A(I) = ArrayY1D(I)
    Next I
    
    '�E�ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Or I = N Then
            ArrayR1D(I, 1) = 0
        Else
            ArrayR1D(I, 1) = 3 * (ArrayY1D(I + 1) - ArrayY1D(I)) / h(I) - 3 * (ArrayY1D(I) - ArrayY1D(I - 1)) / h(I - 1)
        End If
    Next I
    
    '���ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Then
            ArrayL2D(I, 1) = 1
        ElseIf I = N Then
            ArrayL2D(N, N) = 1
        Else
            ArrayL2D(I - 1, I) = h(I - 1)
            ArrayL2D(I, I) = 2 * (h(I) + h(I - 1))
            ArrayL2D(I + 1, I) = h(I)
        End If
    Next I
    
    '���ӂ̔z��̋t�s��
    ArrayLm2D = F_Minverse(ArrayL2D)
    
    'C�̔z������߂�
    C = F_MMult(ArrayLm2D, ArrayR1D)
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


