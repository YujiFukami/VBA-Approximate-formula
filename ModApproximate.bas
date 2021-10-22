Attribute VB_Name = "ModApproximate"
Option Explicit

'TestSplineXY                    �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'TestSpline                      �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'TestSplineXYByArrayX1D          �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'TestSplineByArrayX1D            �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'TestSplineXYPara                �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'TestSplinePara                  �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineXY                        �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineXYByArrayX1D              �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineXYPara                    �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineXYParaFast                �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'Spline                          �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplinePara                      �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineParaFast                  �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'�X�v���C����ԍ������p�ɕ��������E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'ExtractByRangeArray1D           �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineByArrayX1DFast            �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineByArrayX1D                �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineKeisu                     �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'DPH                             �E�E�E���ꏊ�FFukamiAddins3.ModImmediate  
'DebugPrintHairetu               �E�E�E���ꏊ�FFukamiAddins3.ModImmediate  
'��������w��o�C�g���������ɏȗ��E�E�E���ꏊ�FFukamiAddins3.ModImmediate  
'������̊e�����݌v�o�C�g���v�Z  �E�E�E���ꏊ�FFukamiAddins3.ModImmediate  
'�����񕪉�                      �E�E�E���ꏊ�FFukamiAddins3.ModImmediate  
'ExtractArray1D                  �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'CheckArray1D                    �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'CheckArray1DStart1              �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'ExtractArray                    �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'CheckArray2D                    �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'CheckArray2DStart1              �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'UnionArray1D                    �E�E�E���ꏊ�FFukamiAddins3.ModArray      
'F_Minverse                      �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'�����s�񂩃`�F�b�N              �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_MDeterm                       �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Mgyoirekae                    �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Mgyohakidasi                  �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Mjyokyo                       �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_MMult                         �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     



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

Private Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa As Integer, Optional HairetuName As String)
    '20210428�ǉ�
    '���͍������p�ɍ쐬
    
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Private Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa As Integer, Optional HairetuName As String)
'20201023�ǉ�
'20211018 ���͂����z��Hairetu(1 to 1)�̈ꎟ���z��̏ꍇ�ł������ł���悤�ɏC��

    '�񎟌��z����C�~�f�B�G�C�g�E�B���h�E�Ɍ��₷���\������
    
    Dim I       As Long
    Dim J       As Long
    Dim M       As Long
    Dim N       As Long
    Dim TateMin As Long
    Dim TateMax As Long
    Dim YokoMin As Long
    Dim YokoMax As Long

    Dim WithTableHairetu             '�e�[�u���t�z��c�C�~�f�B�G�C�g�E�B���h�E�ɕ\������ۂɃC���f�b�N�X�ԍ���\�������e�[�u����ǉ������z��
    Dim NagasaList
    Dim MaxNagasaList                '�e�����̕����񒷂����i�[�A�e��ł̕����񒷂��̍ő�l���i�[
    Dim NagasaOnajiList              '" "�i���p�X�y�[�X�j�𕶎���ɒǉ����Ċe��ŕ����񒷂��𓯂��ɂ�����������i�[
    Dim OutputList                   '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����镶������i�[
    Const SikiriMoji As String = "|" '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����鎞�Ɋe��̊Ԃɕ\������u�d�؂蕶���v
    
    '������������������������������������������������������
    '���͈����̏���
    Dim Jigen1 As Long
    Dim Jigen2 As Long
    Dim Tmp
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1�����z���2�����z��ɂ���
        Jigen1 = UBound(Hairetu, 1) '20211018 ���͂����z��Hairetu(1 to 1)�̈ꎟ���z��̏ꍇ�ł������ł���悤�ɏC��
        If Jigen1 = 1 Then
            Tmp = Hairetu(Jigen1)
            ReDim Hairetu(1 To 1, 1 To 1)
            Hairetu(1, 1) = Tmp
        Else
            Hairetu = Application.Transpose(Hairetu)
        End If
    End If
    
    TateMin = LBound(Hairetu, 1) '�z��̏c�ԍ��i�C���f�b�N�X�j�̍ŏ�
    TateMax = UBound(Hairetu, 1) '�z��̏c�ԍ��i�C���f�b�N�X�j�̍ő�
    YokoMin = LBound(Hairetu, 2) '�z��̉��ԍ��i�C���f�b�N�X�j�̍ŏ�
    YokoMax = UBound(Hairetu, 2) '�z��̉��ԍ��i�C���f�b�N�X�j�̍ő�
    
    '�e�[�u���t���z��̍쐬
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1) '�e�[�u���ǉ��̕���"+1"����B
    '�uTateMax -TateMin + 1�v�͓��͂����uHairetu�v�̏c�C���f�b�N�X��
    '�uYokoMax -YokoMin + 1�v�͓��͂����uHairetu�v�̉��C���f�b�N�X��
    
    For I = 1 To TateMax - TateMin + 1
        WithTableHairetu(I + 1, 1) = TateMin + I - 1 '�c�e�[�u���iHairetu�̏c�C���f�b�N�X�ԍ��j
        For J = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, J + 1) = YokoMin + J - 1 '���e�[�u���iHairetu�̉��C���f�b�N�X�ԍ��j
            WithTableHairetu(I + 1, J + 1) = Hairetu(I - 1 + TateMin, J - 1 + YokoMin) 'Hairetu�̒��̒l
        Next J
    Next I
    
    '������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\������Ƃ��Ɋe��̕��𓯂��ɐ����邽�߂�
    '�����񒷂��Ƃ��̊e��̍ő�l���v�Z����B
    '�ȉ��ł́uHairetu�v�͈��킸�A�uWithTableHairetu�v�������B
    N = UBound(WithTableHairetu, 1) '�uWithTableHairetu�v�̏c�C���f�b�N�X���i�s���j
    M = UBound(WithTableHairetu, 2) '�uWithTableHairetu�v�̉��C���f�b�N�X���i�񐔁j
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr As String
    For J = 1 To M
        For I = 1 To N
        
            If J > 1 And HyoujiMaxNagasa <> 0 Then
                '�ő�\���������w�肳��Ă���ꍇ�B
                '1��ڂ̃e�[�u���͂��̂܂܂ɂ���B
                TmpStr = WithTableHairetu(I, J)
                WithTableHairetu(I, J) = ��������w��o�C�g���������ɏȗ�(TmpStr, HyoujiMaxNagasa)
            End If
            
            NagasaList(I, J) = LenB(StrConv(WithTableHairetu(I, J), vbFromUnicode)) '�S�p�Ɣ��p����ʂ��Ē������v�Z����B
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����邽�߂�" "(���p�X�y�[�X)��ǉ�����
    '�����񒷂��𓯂��ɂ���B
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa As Long
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) '���̗�̍ő啶���񒷂�
        For I = 1 To N
            'Rept�c�w�蕶������w����A�����ĂȂ�����������o�͂���B
            '�i�ő啶����-�������j�̕�" "�i���p�X�y�[�X�j�����ɂ�������B
            NagasaOnajiList(I, J) = WithTableHairetu(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\�����镶������쐬
    ReDim OutputList(1 To N)
    For I = 1 To N
        For J = 1 To M
            If J = 1 Then
                OutputList(I) = NagasaOnajiList(I, J)
            Else
                OutputList(I) = OutputList(I) & SikiriMoji & NagasaOnajiList(I, J)
            End If
        Next J
    Next I
    
    ''������������������������������������������������������
    '�C�~�f�B�G�C�g�E�B���h�E�ɕ\��
    Debug.Print HairetuName
    For I = 1 To N
        Debug.Print OutputList(I)
    Next I
    
End Sub

Private Function ��������w��o�C�g���������ɏȗ�(Mojiretu As String, ByteNum As Integer)
    '20201023�ǉ�
    '��������w��ȗ��o�C�g�������܂ł̒����ŏȗ�����B
    '�ȗ����ꂽ������̍Ō�̕�����"."�ɕύX����B
    '��FMojiretu = "鳖���" , ByteNum = 6 �c �o�� = "鳖�.."
    '��FMojiretu = "鳖���" , ByteNum = 7 �c �o�� = "鳖��."
    '��FMojiretu = "鳖�XX�" , ByteNum = 6 �c �o�� = "鳖�X."
    '��FMojiretu = "鳖�XX�" , ByteNum = 7 �c �o�� = "鳖�XX."
    
    Dim OriginByte As Integer '���͂���������uMojiretu�v�̃o�C�g������
    Dim Output                '�o�͂���ϐ����i�[
    
    '�uMojiretu�v�̃o�C�g�������v�Z
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    
    If OriginByte <= ByteNum Then
        '�uMojiretu�v�̃o�C�g�������v�Z���ȗ�����o�C�g�������ȉ��Ȃ�
        '�ȗ��͂��Ȃ�
        Output = Mojiretu
    Else
    
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = ������̊e�����݌v�o�C�g���v�Z(Mojiretu)
        BunkaiMojiretu = �����񕪉�(Mojiretu)
        
        Dim AddMoji As String
        AddMoji = "."
        
        Dim I As Long, N As Long
        N = Len(Mojiretu)
        
        For I = 1 To N
            If RuikeiByteList(I) < ByteNum Then
                Output = Output & BunkaiMojiretu(I)
                
            ElseIf RuikeiByteList(I) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(I), vbFromUnicode)) = 1 Then
                    '��FMojiretu = "鳖���" , ByteNum = 6 ,RuikeiByteList(3) = 6
                    'Output = "鳖�.."
                    Output = Output & AddMoji
                Else
                    '��FMojiretu = "鳖�XX�" , ByteNum = 6 ,RuikeiByteList(4) = 6
                    'Output = "鳖�X."
                    Output = Output & AddMoji & AddMoji
                End If
                
                Exit For
                
            ElseIf RuikeiByteList(I) > ByteNum Then
                '��FMojiretu = "鳖���" , ByteNum = 7 ,RuikeiByteList(4) = 8
                'Output = "鳖��."
                Output = Output & AddMoji
                Exit For
            End If
        Next I
        
    End If
        
    ��������w��o�C�g���������ɏȗ� = Output

    
End Function

Private Function ������̊e�����݌v�o�C�g���v�Z(Mojiretu As String)
    '20201023�ǉ�

    '�������1�������ɕ������āA�e�����̃o�C�g���������v�Z���A
    '���̗݌v�l���v�Z����B
    '��FMojiretu="�V�^EK���S��"
    '�o�́�Output = (2,4,5,6,7,10,12)
    
    Dim MojiKosu As Integer
    Dim I        As Long
    Dim TmpMoji  As String
    Dim Output
    MojiKosu = Len(Mojiretu)
    ReDim Output(1 To MojiKosu)
    
    For I = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, I, 1)
        If I = 1 Then
            Output(I) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            Output(I) = LenB(StrConv(TmpMoji, vbFromUnicode)) + Output(I - 1)
        End If
    Next I
    
    ������̊e�����݌v�o�C�g���v�Z = Output
    
End Function

Private Function �����񕪉�(Mojiretu As String)
    '20201023�ǉ�

    '�������1�������������Ĕz��Ɋi�[
    Dim I     As Long
    Dim N     As Long
    Dim Output
    
    N = Len(Mojiretu)
    ReDim Output(1 To N)
    For I = 1 To N
        Output(I) = Mid(Mojiretu, I, 1)
    Next I
    
    �����񕪉� = Output
    
End Function

Private Function ExtractArray1D(Array1D, StartNum As Long, EndNum As Long)
'�ꎟ���z��̎w��͈͂�z��Ƃ��Ē��o����
'20211009

'����
'Array1D �E�E�E�ꎟ���z��
'StartNum�E�E�E���o�͈͂̊J�n�ԍ�
'EndNum  �E�E�E���o�͈͂̏I���ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I As Long
    Dim N As Long
    N = UBound(Array1D, 1) '�v�f��
    
    If StartNum > EndNum Then
        MsgBox ("���o�͈͂̊J�n�ʒu�uStartNum�v�́A�I���ʒu�uEndNum�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        Exit Function
    ElseIf StartNum < 1 Then
        MsgBox ("���o�͈͂̊J�n�ʒu�uStartNum�v��1�ȏ�̒l�����Ă�������")
        Stop
        Exit Function
    ElseIf EndNum > N Then
        MsgBox ("���o�͈͂̏I���s�uEndNum�v�͒��o���̈ꎟ���z��̗v�f��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        Exit Function
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndNum - StartNum + 1)
    
    For I = StartNum To EndNum
        Output(I - StartNum + 1) = Array1D(I)
    Next I
    
    '�o��
    ExtractArray1D = Output
    
End Function

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function ExtractArray(Array2D, StartRow As Long, StartCol As Long, EndRow As Long, EndCol As Long)
'�񎟌��z��̎w��͈͂�z��Ƃ��Ē��o����
'20210917

'����
'Array2D �E�E�E�񎟌��z��
'StartRow�E�E�E���o�͈͂̊J�n�s�ԍ�
'StartCol�E�E�E���o�͈͂̊J�n��ԍ�
'EndRow  �E�E�E���o�͈͂̏I���s�ԍ�
'EndCol  �E�E�E���o�͈͂̏I����ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
    
    If StartRow > EndRow Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v�́A�I���s�uEndRow�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v�́A�I����uEndCol�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("���o�͈͂̏I���s�uStartRow�v�͒��o���̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("���o�͈͂̏I����uStartCol�v�͒��o���̓񎟌��z��̗�" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '�o��
    ExtractArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function UnionArray1D(UpperArray1D, LowerArray1D)
'�ꎟ���z�񓯎m����������1�̔z��Ƃ���B
'20210923

'UpperArray1D�E�E�E��Ɍ�������ꎟ���z��
'LowerArray1D�E�E�E���Ɍ�������ꎟ���z��

    '�����`�F�b�N
    Call CheckArray1D(UpperArray1D, "UpperArray1D")
    Call CheckArray1DStart1(UpperArray1D, "UpperArray1D")
    Call CheckArray1D(LowerArray1D, "LowerArray1D")
    Call CheckArray1DStart1(LowerArray1D, "LowerArray1D")
    
    '����
    Dim I  As Long
    Dim N1 As Long
    Dim N2 As Long
    N1 = UBound(UpperArray1D, 1)
    N2 = UBound(LowerArray1D, 1)
    Dim Output
    ReDim Output(1 To N1 + N2)
    For I = 1 To N1
        Output(I) = UpperArray1D(I)
    Next I
    For I = 1 To N2
        Output(N1 + I) = LowerArray1D(I)
    Next I
    
    '�o��
    UnionArray1D = Output
    
End Function

Private Function F_Minverse(ByVal Matrix)
    '20210603����
    'F_Minverse(input_M)
    'F_Minverse(�z��)
    '�]���q�s���p���ċt�s����v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Matrix)
    
    '�v�Z����������������������������������������������������������
    Dim I        As Integer
    Dim J        As Integer
    Dim N        As Integer
    Dim Output() As Double
    N = UBound(Matrix, 1)
    ReDim Output(1 To N, 1 To N)
    
    Dim detM As Double '�s�񎮂̒l���i�[
    detM = F_MDeterm(Matrix) '�s�񎮂����߂�
    
    Dim Mjyokyo '�w��̗�E�s�����������z����i�[
    
    For I = 1 To N '�e��
        For J = 1 To N '�e�s
            
            'I��,J�s����������
            Mjyokyo = F_Mjyokyo(Matrix, J, I)
            
            'I��,J�s�̗]���q�����߂ďo�͂���t�s��Ɋi�[
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_Minverse = Output
    
End Function

Private Sub �����s�񂩃`�F�b�N(Matrix)
    '20210603�ǉ�
    
    If UBound(Matrix, 1) <> UBound(Matrix, 2) Then
        MsgBox ("�����s�����͂��Ă�������" & vbLf & _
                "���͂��ꂽ�z��̗v�f����" & "�u" & _
                UBound(Matrix, 1) & "�~" & UBound(Matrix, 2) & "�v" & "�ł�")
        Stop
        End
    End If

End Sub

Private Function F_MDeterm(Matrix)
    '20210603����
    'F_MDeterm(Matrix)
    'F_MDeterm(�z��)
    '�s�񎮂��v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Matrix)
    
    '�v�Z����������������������������������������������������������
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    N = UBound(Matrix, 1)
    
    Dim Matrix2 '�|���o�����s���s��
    Matrix2 = Matrix
    
    For I = 1 To N '�e��
        For J = I To N '�|���o�����̍s�̒T��
            If Matrix2(J, I) <> 0 Then
                K = J '�|���o�����̍s
                Exit For
            End If
            
            If J = N And Matrix2(J, I) = 0 Then '�|���o�����̒l���S��0�Ȃ�s�񎮂̒l��0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I��,I�s)�ȊO�ő|���o���ƂȂ�ꍇ�͍s�����ւ�
            Matrix2 = F_Mgyoirekae(Matrix2, I, K)
        End If
        
        '�|���o��
        Matrix2 = F_Mgyohakidasi(Matrix2, I, I)
              
    Next I
    
    
    '�s�񎮂̌v�Z
    Dim Output As Double
    Output = 1
    
    For I = 1 To N '�e(I��,I�s)���|�����킹�Ă���
        Output = Output * Matrix2(I, I)
    Next I
    
    '�o�́�����������������������������������������������������
    F_MDeterm = Output
    
End Function

Private Function F_Mgyoirekae(Matrix, Row1 As Integer, Row2 As Integer)
    '20210603����
    'F_Mgyoirekae(Matrix, Row1, Row2)
    'F_Mgyoirekae(�z��,�w��s�ԍ��@,�w��s�ԍ��A)
    '�s��Matrix�̇@�s�ƇA�s�����ւ���
    
    Dim I     As Integer
    Dim J     As Integer
    Dim K     As Integer
    Dim M     As Integer
    Dim N     As Integer
    Dim Output
    
    Output = Matrix
    M = UBound(Matrix, 2) '�񐔎擾
    
    For I = 1 To M
        Output(Row2, I) = Matrix(Row1, I)
        Output(Row1, I) = Matrix(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Private Function F_Mgyohakidasi(Matrix, Row As Integer, Col As Integer)
    '20210603����
    'F_Mgyohakidasi(Matrix, Row, Col)
    'F_Mgyohakidasi(�z��,�w��s,�w���)
    '�s��Matrix��Row�s�Col��̒l�Ŋe�s��|���o��
    
    Dim I     As Integer
    Dim J     As Integer
    Dim N     As Integer
    Dim Output
    
    Output = Matrix
    N = UBound(Output, 1) '�s���擾
    
    Dim Hakidasi '�|���o�����̍s
    Dim X As Double '�|���o�����̒l
    Dim Y As Double
    ReDim Hakidasi(1 To N)
    X = Matrix(Row, Col)
    
    For I = 1 To N '�|���o������1�s���쐬
        Hakidasi(I) = Matrix(Row, I)
    Next I
    
    For I = 1 To N '�e�s
        If I = Row Then
            '�|���o�����̍s�̏ꍇ�͂��̂܂�
            For J = 1 To N
                Output(I, J) = Matrix(I, J)
            Next J
        
        Else
            '�|���o�����̍s�ȊO�̏ꍇ�͑|���o��
            Y = Matrix(I, Col) '�|���o����̗�̒l
            For J = 1 To N
                Output(I, J) = Matrix(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Private Function F_Mjyokyo(Matrix, Row As Integer, Col As Integer)
    '20210603����
    'F_Mjyokyo(Matrix, Row, Col)
    'F_Mjyokyo(�z��,�w��s,�w���)
    '�s��Matrix��Row�s�ACol������������s���Ԃ�
    
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim M As Integer
    Dim N As Integer '�����グ�p(Integer�^)
    Dim Output '�w�肵���s�E���������̔z��
    
    N = UBound(Matrix, 1) '�s���擾
    M = UBound(Matrix, 2) '�񐔎擾
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2 As Integer
    Dim J2 As Integer
    
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
                    Output(I2, J2) = Matrix(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function

Private Function F_MMult(ByVal Matrix1, ByVal Matrix2)
    'F_MMult(Matrix1, Matrix2)
    'F_MMult(�z��@,�z��A)
    '�s��̐ς��v�Z
    '20180213����
    '20210603����
    
    '���͒l�̃`�F�b�N�ƏC��������������������������������������������������������
    '�z��̎����`�F�b�N
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(Matrix1, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(Matrix2, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����1�Ȃ玟��2�ɂ���B��)�z��(1 to N)���z��(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Matrix1 = Application.Transpose(Matrix1)
    End If
    If IsEmpty(JigenCheck2) Then
        Matrix2 = Application.Transpose(Matrix2)
    End If
    
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If UBound(Matrix1, 1) = 0 Or UBound(Matrix1, 2) = 0 Then
        Matrix1 = Application.Transpose(Application.Transpose(Matrix1))
    End If
    If UBound(Matrix2, 1) = 0 Or UBound(Matrix2, 2) = 0 Then
        Matrix2 = Application.Transpose(Application.Transpose(Matrix2))
    End If
    
    '���͒l�̃`�F�b�N
    If UBound(Matrix1, 2) <> UBound(Matrix2, 1) Then
        MsgBox ("�z��1�̗񐔂Ɣz��2�̍s������v���܂���B" & vbLf & _
               "(�o��) = (�z��1)(�z��2)")
        Stop
        End
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I        As Integer
    Dim J        As Integer
    Dim K        As Integer
    Dim M        As Integer
    Dim N        As Integer
    Dim M2       As Integer
    Dim Output() As Double '�o�͂���z��
    N = UBound(Matrix1, 1) '�z��1�̍s��
    M = UBound(Matrix1, 2) '�z��1�̗�
    M2 = UBound(Matrix2, 2) '�z��2�̗�
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '�e�s
        For J = 1 To M2 '�e��
            For K = 1 To M '(�z��1��I�s)��(�z��2��J��)���|�����킹��
                Output(I, J) = Output(I, J) + Matrix1(I, K) * Matrix2(K, J)
            Next K
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_MMult = Output
    
End Function


