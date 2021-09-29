Attribute VB_Name = "ModApproximate"
'シート関数用近似、補間関数
Private Sub TestSplineXY()
'SplineXYの実行テスト

    Dim ArrayXY2D, InputX#, OutputY#
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

    Dim ArrayX1D, ArrayY1D, InputX#
    ArrayX1D = Application.Transpose(Application.Transpose( _
                Array(0, 1, 2, 3, 4, 5, 6) _
                ))
    
    ArrayY1D = Application.Transpose(Application.Transpose( _
                Array(5.93255769237665, 9.99308268536779, 5.5044328013839, 5.60877013928983, 1.51682123665907, 8.18738634627902, 4.42233332813268) _
                ))
                
    InputX = 3.5
    
    Dim OutputY#
    OutputY = Spline(ArrayX1D, ArrayY1D, InputX)

    Debug.Print "OutputY = " & OutputY
    
End Sub

Private Sub TestSplineXYByArrayX1D()
'SplineXYByArrayX1Dの実行テスト

    Dim ArrayXY2D, InputArrayX1D
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

    Dim ArrayX1D, ArrayY1D, InputArrayX1D
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

    Dim ArrayXY2D, BunkatuN&
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

    Dim ArrayX1D, ArrayY1D, BunkatuN&
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

Function SplineXY(ByVal ArrayXY2D, InputX#)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力値Xに対する補間値Y
    
    '＜入力値の説明＞
    'ArrayXY2D：補間の対象となるX,Yの値が格納された配列
    'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
    'X:補間位置のXの値
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D, ArrayY1D
    Dim I%, N%
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(1 To N)
    ReDim ArrayY1D(1 To N)
    
    For I = 1 To N
        ArrayX1D(I) = ArrayXY2D(I, 1)
        ArrayY1D(I) = ArrayXY2D(I, 2)
    Next I
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim OutputY#
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
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
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
    
    Dim ArrayX1D, ArrayY1D
    Dim I%, N%
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

Function SplineXYPara(ByVal ArrayXY2D, BunkatuN&)
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
    Dim StartNum%
    StartNum = LBound(ArrayXY2D) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D, ArrayY1D
    Dim I%, N%
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(StartNum To StartNum - 1 + N)
    ReDim ArrayY1D(StartNum To StartNum - 1 + N)
    
    For I = 1 To N
        ArrayX1D(I + StartNum - 1) = ArrayXY2D(I, 1)
        ArrayY1D(I + StartNum - 1) = ArrayXY2D(I, 2)
    Next I
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim Dummy
    Dim OutputArrayX1D, OutputArrayY1D
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

Function Spline#(ByVal ArrayX1D, ByVal ArrayY1D, InputX#)
        
    '20171124修正
    '20180309改良
    
    'スプライン補間計算を行う
    '
    '<出力値の説明>
    '入力値Xに対する補間値Y
    '
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
    Dim JigenCheck1%, JigenCheck2%
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
    Dim I%, N%, K%, A, B, C, D
    Dim OutputY# '出力値Y
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

Function SplinePara(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN&)
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
    
    Dim StartNum%
    '行列の開始要素を1に変更（計算しやすいから）
    StartNum = LBound(ArrayX1D, 1) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%
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
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D#(), ArrayParaT1D#()
    
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
    
    Dim OutputArrayX1D, OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1D(ArrayT1D, ArrayX1D, ArrayParaT1D)
    OutputArrayY1D = SplineByArrayX1D(ArrayT1D, ArrayY1D, ArrayParaT1D)
    
    '出力
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplinePara = Output
    
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
    
    Dim StartNum%
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
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
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
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(ArrayX1D, 1) '補間対象の要素数
    
    Dim OutputArrayY1D#() '出力するYの格納
    Dim NX%
    NX = UBound(InputArrayX1D, 1) '補間位置の個数
    ReDim OutputArrayY1D(1 To NX)
    Dim TmpX#, TmpY#
    
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
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim A, B, C, D
    N = UBound(ArrayX1D, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h#()
    Dim ArrayL2D#() '左辺の配列 要素数(1 to N,1 to N)
    Dim ArrayR1D#() '右辺の配列 要素数(1 to N,1 to 1)
    Dim ArrayLm2D#() '左辺の配列の逆行列 要素数(1 to N,1 to N)
    
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


