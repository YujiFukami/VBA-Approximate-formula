Attribute VB_Name = "ModApproximate"
'シート関数用近似、補間関数
Private Sub TestSplineXY()
'SplineXYの実行テスト

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
'Splineの実行テスト

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
'SplineXYByArrayX1Dの実行テスト

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
'SplineByArrayX1Dの実行テスト

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
'SplineXYParaの実行テスト

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
'SplineParaの実行テスト

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
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力値Xに対する補間値Y
    
    '＜入力値の説明＞
    'ArrayXY2D：補間の対象となるX,Yの値が格納された配列
    'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
    'X:補間位置のXの値
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    Dim RangeNaraTrue As Boolean: RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim OutputY As Double
    OutputY = Spline(ArrayX1D, ArrayY1D, InputX)
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    If RangeNaraTrue Then
        'ワークシート関数の場合
        SplineXY = Application.Transpose(OutputY)
    Else
        'VBA上での処理の場合
        SplineXY = OutputY
    End If
    
End Function

Function SplineXYByArrayX1D(ByVal ArrayXY2D, ByVal InputArrayX1D)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力配列InputArrayX1Dに対する補間値の配列YList
    
    '＜入力値の説明＞
    'HariretuXY：補間の対象となるX,Yの値が格納された配列
    'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
    'InputArrayX1D:補間位置Xが格納された配列
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    Dim RangeNaraTrue As Boolean: RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If

    '行列の開始要素を1に変更（計算しやすいから）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim OutputArrayY1D
    OutputArrayY1D = SplineByArrayX1D(ArrayX1D, ArrayY1D, InputArrayX1D)
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    If RangeNaraTrue = True Then
        'ワークシート関数の場合
        SplineXYByArrayX1D = Application.Transpose(OutputArrayY1D)
    Else
        'VBA上での処理の場合
        SplineXYByArrayX1D = OutputArrayY1D
    End If
    
End Function

Function SplineXYPara(ByVal ArrayXY2D, BunkatuN As Long)
    'パラメトリック関数形式でスプライン補間を行う
    'ArrayX,ArrayYがどちらも単調増加、単調減少でない場合に用いる。
    '＜出力値の説明＞
    'パラメトリック関数形式で補間されたXList,YListが格納されたXYList
    '1列目がXList,2列目がYList
    
    '＜入力値の説明＞
    'ArrayXY2D：補間の対象となるX,Yの値が格納された配列
    'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
    'パラメトリック関数の分割個数（出力されるXList,YListの要素数は(分割個数+1)）
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
    End If
        
    '行列の開始要素を1に変更（計算しやすいから）
    Dim StartNum As Integer
    StartNum = LBound(ArrayXY2D) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
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
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    SplineXYPara = OutputArrayXY2D
    
End Function

Function SplineXYParaFast(ByVal ArrayXY2D, BunkatuN As Long, PointCount As Long)
'パラメトリック関数形式でスプライン補間を行う
'分割して計算を高速化する
'ArrayX,ArrayYがどちらも単調増加、単調減少でない場合に用いる。
    
'引数
'ArrayXY2D ・・・補間の対象となるX,Yの値が格納された配列
'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
'BunkatuN  ・・・パラメトリック関数の分割個数（出力されるXList,YListの要素数は(分割個数+1)）
'PointCount・・・分割する際の点数
    
'返り値
'パラメトリック関数形式で補間されたXList,YListが格納されたXYList
'1列目がXList,2列目がYList
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
    End If
        
    '行列の開始要素を1に変更（計算しやすいから）
    Dim StartNum As Integer
    StartNum = LBound(ArrayXY2D) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
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
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    SplineXYParaFast = OutputArrayXY2D
    
End Function


Function Spline(ByVal ArrayX1D, ByVal ArrayY1D, InputX As Double)
        
    '20171124修正
    '20180309改良
    
    'スプライン補間計算を行う
    
    '<出力値の説明>
    '入力値Xに対する補間値Y
    
    '<入力値の説明>
    'ArrayX1D：補間の対象とする配列X
    'ArrayY1D：補間の対象とする配列Y
    'InputX  ：補間位置のXの値
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I As Integer
    Dim N As Integer
    Dim K As Integer
    Dim A
    Dim B
    Dim C
    Dim D
    Dim OutputY As Double '出力値Y
    Dim SotoNaraTrue As Boolean
    SotoNaraTrue = False
    N = UBound(ArrayX1D, 1)
       
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
        
    For I = 1 To N - 1
        If ArrayX1D(I) < ArrayX1D(I + 1) Then 'Xが単調増加の場合
            If I = 1 And ArrayX1D(1) > InputX Then '範囲に入らないとき(開始点より前)
                OutputY = ArrayY1D(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And ArrayX1D(I + 1) <= InputX Then '範囲に入らないとき(終了点より後)
                OutputY = ArrayY1D(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf ArrayX1D(I) <= InputX And ArrayX1D(I + 1) > InputX Then '範囲内
                K = I: Exit For
            
            End If
        Else 'Xが単調減少の場合
        
            If I = 1 And ArrayX1D(1) < InputX Then '範囲に入らないとき(開始点より前)
                OutputY = ArrayY1D(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And ArrayX1D(I + 1) >= InputX Then '範囲に入らないとき(終了点より後)
                OutputY = ArrayY1D(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf ArrayX1D(I + 1) < InputX And ArrayX1D(I) >= InputX Then '範囲内
                K = I: Exit For
            
            End If
        
        End If
    Next I
        
    If SotoNaraTrue = False Then
        OutputY = A(K) + B(K) * (InputX - ArrayX1D(K)) + C(K) * (InputX - ArrayX1D(K)) ^ 2 + D(K) * (InputX - ArrayX1D(K)) ^ 3
    End If
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Spline = OutputY

End Function

Function SplinePara(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN As Long)
    'パラメトリック関数形式でスプライン補間を行う
    'ArrayX1D,ArrayY1Dがどちらも単調増加、単調減少でない場合に用いる。
    '＜出力値の説明＞
    'パラメトリック関数形式で補間されたXList,YList
    
    '＜入力値の説明＞
    'ArrayX1D：補間の対象となるXの値が格納された配列
    'ArrayY1D：補間の対象となるYの値が格納された配列
    'パラメトリック関数の分割個数（出力されるOutputArrayX1D,OutputArrayY1Dの要素数は(分割個数+1)）
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    Dim StartNum As Integer
    '行列の開始要素を1に変更（計算しやすいから）
    StartNum = LBound(ArrayX1D, 1) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D()     As Double
    Dim ArrayParaT1D() As Double
    
    'X,Yの補間の基準となる配列を作成
    ReDim ArrayT1D(1 To N)
    For I = 1 To N
        '0～1を等間隔
        ArrayT1D(I) = (I - 1) / (N - 1)
    Next I
    
    '出力補間位置の基準位置
    If JigenCheck1 > 0 Then '出力値の形状を入力値に合わせるための処理
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            ArrayParaT1D(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            ArrayParaT1D(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim OutputArrayX1D
    Dim OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1D(ArrayT1D, ArrayX1D, ArrayParaT1D)
    OutputArrayY1D = SplineByArrayX1D(ArrayT1D, ArrayY1D, ArrayParaT1D)
    
    '出力
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplinePara = Output
    
End Function

Function SplineParaFast(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN As Long, PointCount As Long)
'パラメトリック関数形式でスプライン補間を行う
'分割して計算を高速化する
'ArrayX1D,ArrayY1Dがどちらも単調増加、単調減少でない場合に用いる。
'20211009

'引数
'ArrayX1D  ・・・補間の対象となるXの値が格納された配列
'ArrayY1D  ・・・補間の対象となるYの値が格納された配列
'BunkatuN  ・・・パラメトリック関数の分割個数（出力されるOutputArrayX1D,OutputArrayY1Dの要素数は(分割個数+1)）
'PointCount・・・分割する際の点数

'返り値
'パラメトリック関数形式で補間されたXList,YList
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    Dim StartNum As Integer
    '行列の開始要素を1に変更（計算しやすいから）
    StartNum = LBound(ArrayX1D, 1) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I As Integer
    Dim N As Integer
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D() As Double, ArrayParaT1D() As Double
    
    'X,Yの補間の基準となる配列を作成
    ReDim ArrayT1D(1 To N)
    For I = 1 To N
        '0～1を等間隔
        ArrayT1D(I) = (I - 1) / (N - 1)
    Next I
    
    '出力補間位置の基準位置
    If JigenCheck1 > 0 Then '出力値の形状を入力値に合わせるための処理
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            ArrayParaT1D(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            ArrayParaT1D(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim OutputArrayX1D
    Dim OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1DFast(ArrayT1D, ArrayX1D, ArrayParaT1D, PointCount)
    OutputArrayY1D = SplineByArrayX1DFast(ArrayT1D, ArrayY1D, ArrayParaT1D, PointCount)
    
    '出力
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplineParaFast = Output
    
End Function


Private Function スプライン補間高速化用に分割処理(ByVal ArrayX1D, ByVal ArrayY1D, ByVal CalPoint1D, PointCount As Long)
'スプライン補間高速化用に分割処理
'20211009

'引数
'ArrayX1D  ・・・補間元のX座標リスト
'ArrayY1D  ・・・補間元のY座標リスト
'CalPoint1D・・・補間位置のX座標リスト
'PointCount・・・分割後の一つの分割の点数

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
    
    Dim Output '出力値格納変数
    ReDim Output(1 To N, 1 To 3) '1:補間元X座標リスト,2:補間元Y座標リスト,3:補間位置X座標リスト
    'Nはとりあえずの最大で、後で配列を縮小する
    
    Dim TmpXList
    Dim TmpYList
    Dim TmpPointList
    Dim TmpInterXList
    Dim StartNum      As Long '分割する補間元座標の開始位置
    Dim EndNum        As Long '分割する補間元座標の終了位置
    Dim InterStartNum As Long '分割された補間元座標で実際の補間範囲の開始位置
    Dim InterEndNum   As Long '分割された補間元座標で実際の補間範囲の終了位置
    
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
    
    '出力する格納配列の範囲調整
    Output = ExtractArray(Output, 1, 1, K, 3)
    
    '分割した補間位置で重複するものを消去
    N = UBound(Output, 1)
    Dim TmpList1
    Dim TmpList2
    For I = 2 To N
        TmpList1 = Output(I - 1, 3)
        TmpList2 = Output(I, 3)
        If IsEmpty(TmpList1) = False And IsEmpty(TmpList2) = False Then
            If TmpList1(UBound(TmpList1, 1)) = TmpList2(1) Then '最後の要素と最初の要素を比較する
                If UBound(TmpList2, 1) = 1 Then
                    TmpList2 = Empty
                Else
                    TmpList2 = ExtractArray1D(TmpList2, 2, UBound(TmpList2, 1))
                End If
                Output(I, 3) = TmpList2
            End If
        End If
    Next
    
    スプライン補間高速化用に分割処理 = Output
    
End Function

Function ExtractByRangeArray1D(InputArray1D, RangeArray1D)
'一次元配列の指定範囲を抽出する。
'指定範囲はRangeArray1Dで指定する。
'20211009

'引数
'InputArray1D・・・抽出元の一次元配列
'RangeArray1D・・・抽出する範囲を指定する一次元配列

'例
'InputArray1D = (1,2,3,4,5,6,7,8,9,10)
'RangeArray1D = (3,4,7)
'出力 = (3,4,5,6,7)

    '引数チェック
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
    
    '指定範囲の最小、最大を取得
    Dim MinNum As Double
    Dim MaxNum As Double
    MinNum = WorksheetFunction.Min(RangeArray1D)
    MaxNum = WorksheetFunction.Max(RangeArray1D)
    
    '抽出範囲の開始位置、終了位置を計算
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
        '抽出範囲なしでEmptyを返す
        Exit Function
    End If
    
    For I = StartNum To N
        If InputArray1D(I) > MaxNum Then
            EndNum = I - 1
            Exit For
        End If
    Next
    
    If EndNum = 0 Then
        '終了位置が見つからない場合は終了まで全部含む
        EndNum = N
    End If
    
    '範囲抽出
    Dim Output '出力値格納変数
    Output = ExtractArray1D(InputArray1D, StartNum, EndNum)
    
    '出力
    ExtractByRangeArray1D = Output
    
End Function

Function SplineByArrayX1DFast(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D, PointCount As Long)
 'スプライン補間計算を行う
 '分割して計算することで高速化する

'引数
'HairetuX     ・・・補間の対象となるXの値が格納された配列
'HairetuY     ・・・補間の対象となるYの値が格納された配列
'InputArrayX1D・・・補間位置Xが格納された配列
'PointCount   ・・・分割する際の点数

'返り値
'入力配列InputArrayX1Dに対する補間値の配列
        
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
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
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1Dの開始要素番号を取っておく（出力値を合わせるため）
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    Dim JigenCheck3 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck3 = UBound(InputArrayX1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim SplitArrayList
    SplitArrayList = スプライン補間高速化用に分割処理(ArrayX1D, ArrayY1D, InputArrayX1D, PointCount)
        
    Dim TmpXList
    Dim TmpYList
    Dim TmpPointList
    Dim Output '出力値格納変数
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
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力配列InputArrayX1Dに対する補間値の配列YList
    
    '＜入力値の説明＞
    'HairetuX：補間の対象となるXの値が格納された配列
    'HairetuY：補間の対象となるYの値が格納された配列
    'InputArrayX1D:補間位置Xが格納された配列

    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
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
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1Dの開始要素番号を取っておく（出力値を合わせるため）
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    Dim JigenCheck3 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck3 = UBound(InputArrayX1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim A, B, C, D
    Dim I As Long, J As Long, K As Long, M As Long, N As Long '数え上げ用(Long型)
    
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(ArrayX1D, 1) '補間対象の要素数
    
    Dim OutputArrayY1D() As Double '出力するYの格納
    Dim NX As Integer
    NX = UBound(InputArrayX1D, 1) '補間位置の個数
    ReDim OutputArrayY1D(1 To NX)
    Dim TmpX As Double, TmpY As Double
    
    For J = 1 To NX
        TmpX = InputArrayX1D(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If ArrayX1D(I) < ArrayX1D(I + 1) Then 'Xが単調増加の場合
                If I = 1 And ArrayX1D(1) > TmpX Then '範囲に入らないとき(開始点より前)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) <= TmpX Then '範囲に入らないとき(終了点より後)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I) <= TmpX And ArrayX1D(I + 1) > TmpX Then '範囲内
                    K = I: Exit For
                
                End If
            Else 'Xが単調減少の場合
            
                If I = 1 And ArrayX1D(1) < TmpX Then '範囲に入らないとき(開始点より前)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) >= TmpX Then '範囲に入らないとき(終了点より後)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I + 1) < TmpX And ArrayX1D(I) >= TmpX Then '範囲内
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - ArrayX1D(K)) + C(K) * (TmpX - ArrayX1D(K)) ^ 2 + D(K) * (TmpX - ArrayX1D(K)) ^ 3
        End If
        
        OutputArrayY1D(J) = TmpY
        
    Next J
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim Output
    
    '出力する配列を入力した配列InputArrayX1Dの形状に合わせる
    If JigenCheck3 = 1 Then '入力のInputArrayX1Dが二次元配列
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
        'ワークシート関数の場合
        SplineByArrayX1D = Application.Transpose(Output)
    Else
        'VBA上での処理の場合
        SplineByArrayX1D = Output
    End If
    
End Function

Function SplineKeisu(ByVal ArrayX1D, ByVal ArrayY1D)

    '参考：http://www5d.biglobe.ne.jp/stssk/maze/spline.html
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
    Dim ArrayL2D()  As Double '左辺の配列 要素数(1 to N,1 to N)
    Dim ArrayR1D()  As Double '右辺の配列 要素数(1 to N,1 to 1)
    Dim ArrayLm2D() As Double '左辺の配列の逆行列 要素数(1 to N,1 to N)
    
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
    
    '右辺の配列の計算
    For I = 1 To N
        If I = 1 Or I = N Then
            ArrayR1D(I, 1) = 0
        Else
            ArrayR1D(I, 1) = 3 * (ArrayY1D(I + 1) - ArrayY1D(I)) / h(I) - 3 * (ArrayY1D(I) - ArrayY1D(I - 1)) / h(I - 1)
        End If
    Next I
    
    '左辺の配列の計算
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
    
    '左辺の配列の逆行列
    ArrayLm2D = F_Minverse(ArrayL2D)
    
    'Cの配列を求める
    C = F_MMult(ArrayLm2D, ArrayR1D)
    C = Application.Transpose(C)
    
    'Bの配列を求める
    For I = 1 To N - 1
        B(I) = (A(I + 1) - A(I)) / h(I) - h(I) * (C(I + 1) + 2 * C(I)) / 3
    Next I
    
    'Dの配列を求める
    For I = 1 To N - 1
        D(I) = (C(I + 1) - C(I)) / (3 * h(I))
    Next I
    
    '出力
    Dim Output(1 To 4)
    Output(1) = A
    Output(2) = B
    Output(3) = C
    Output(4) = D
    
    SplineKeisu = Output

End Function


