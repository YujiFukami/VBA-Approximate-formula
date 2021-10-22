Attribute VB_Name = "ModApproximate"
Option Explicit

'TestSplineXY                    ・・・元場所：FukamiAddins3.ModApproximate
'TestSpline                      ・・・元場所：FukamiAddins3.ModApproximate
'TestSplineXYByArrayX1D          ・・・元場所：FukamiAddins3.ModApproximate
'TestSplineByArrayX1D            ・・・元場所：FukamiAddins3.ModApproximate
'TestSplineXYPara                ・・・元場所：FukamiAddins3.ModApproximate
'TestSplinePara                  ・・・元場所：FukamiAddins3.ModApproximate
'SplineXY                        ・・・元場所：FukamiAddins3.ModApproximate
'SplineXYByArrayX1D              ・・・元場所：FukamiAddins3.ModApproximate
'SplineXYPara                    ・・・元場所：FukamiAddins3.ModApproximate
'SplineXYParaFast                ・・・元場所：FukamiAddins3.ModApproximate
'Spline                          ・・・元場所：FukamiAddins3.ModApproximate
'SplinePara                      ・・・元場所：FukamiAddins3.ModApproximate
'SplineParaFast                  ・・・元場所：FukamiAddins3.ModApproximate
'スプライン補間高速化用に分割処理・・・元場所：FukamiAddins3.ModApproximate
'ExtractByRangeArray1D           ・・・元場所：FukamiAddins3.ModApproximate
'SplineByArrayX1DFast            ・・・元場所：FukamiAddins3.ModApproximate
'SplineByArrayX1D                ・・・元場所：FukamiAddins3.ModApproximate
'SplineKeisu                     ・・・元場所：FukamiAddins3.ModApproximate
'DPH                             ・・・元場所：FukamiAddins3.ModImmediate  
'DebugPrintHairetu               ・・・元場所：FukamiAddins3.ModImmediate  
'文字列を指定バイト数文字数に省略・・・元場所：FukamiAddins3.ModImmediate  
'文字列の各文字累計バイト数計算  ・・・元場所：FukamiAddins3.ModImmediate  
'文字列分解                      ・・・元場所：FukamiAddins3.ModImmediate  
'ExtractArray1D                  ・・・元場所：FukamiAddins3.ModArray      
'CheckArray1D                    ・・・元場所：FukamiAddins3.ModArray      
'CheckArray1DStart1              ・・・元場所：FukamiAddins3.ModArray      
'ExtractArray                    ・・・元場所：FukamiAddins3.ModArray      
'CheckArray2D                    ・・・元場所：FukamiAddins3.ModArray      
'CheckArray2DStart1              ・・・元場所：FukamiAddins3.ModArray      
'UnionArray1D                    ・・・元場所：FukamiAddins3.ModArray      
'F_Minverse                      ・・・元場所：FukamiAddins3.ModMatrix     
'正方行列かチェック              ・・・元場所：FukamiAddins3.ModMatrix     
'F_MDeterm                       ・・・元場所：FukamiAddins3.ModMatrix     
'F_Mgyoirekae                    ・・・元場所：FukamiAddins3.ModMatrix     
'F_Mgyohakidasi                  ・・・元場所：FukamiAddins3.ModMatrix     
'F_Mjyokyo                       ・・・元場所：FukamiAddins3.ModMatrix     
'F_MMult                         ・・・元場所：FukamiAddins3.ModMatrix     



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

Private Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa As Integer, Optional HairetuName As String)
    '20210428追加
    '入力高速化用に作成
    
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Private Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa As Integer, Optional HairetuName As String)
'20201023追加
'20211018 入力した配列がHairetu(1 to 1)の一次元配列の場合でも処理できるように修正

    '二次元配列をイミディエイトウィンドウに見やすく表示する
    
    Dim I       As Long
    Dim J       As Long
    Dim M       As Long
    Dim N       As Long
    Dim TateMin As Long
    Dim TateMax As Long
    Dim YokoMin As Long
    Dim YokoMax As Long

    Dim WithTableHairetu             'テーブル付配列…イミディエイトウィンドウに表示する際にインデックス番号を表示したテーブルを追加した配列
    Dim NagasaList
    Dim MaxNagasaList                '各文字の文字列長さを格納、各列での文字列長さの最大値を格納
    Dim NagasaOnajiList              '" "（半角スペース）を文字列に追加して各列で文字列長さを同じにした文字列を格納
    Dim OutputList                   'イミディエイトウィンドウに表示する文字列を格納
    Const SikiriMoji As String = "|" 'イミディエイトウィンドウに表示する時に各列の間に表示する「仕切り文字」
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力引数の処理
    Dim Jigen1 As Long
    Dim Jigen2 As Long
    Dim Tmp
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1次元配列は2次元配列にする
        Jigen1 = UBound(Hairetu, 1) '20211018 入力した配列がHairetu(1 to 1)の一次元配列の場合でも処理できるように修正
        If Jigen1 = 1 Then
            Tmp = Hairetu(Jigen1)
            ReDim Hairetu(1 To 1, 1 To 1)
            Hairetu(1, 1) = Tmp
        Else
            Hairetu = Application.Transpose(Hairetu)
        End If
    End If
    
    TateMin = LBound(Hairetu, 1) '配列の縦番号（インデックス）の最小
    TateMax = UBound(Hairetu, 1) '配列の縦番号（インデックス）の最大
    YokoMin = LBound(Hairetu, 2) '配列の横番号（インデックス）の最小
    YokoMax = UBound(Hairetu, 2) '配列の横番号（インデックス）の最大
    
    'テーブル付き配列の作成
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1) 'テーブル追加の分で"+1"する。
    '「TateMax -TateMin + 1」は入力した「Hairetu」の縦インデックス数
    '「YokoMax -YokoMin + 1」は入力した「Hairetu」の横インデックス数
    
    For I = 1 To TateMax - TateMin + 1
        WithTableHairetu(I + 1, 1) = TateMin + I - 1 '縦テーブル（Hairetuの縦インデックス番号）
        For J = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, J + 1) = YokoMin + J - 1 '横テーブル（Hairetuの横インデックス番号）
            WithTableHairetu(I + 1, J + 1) = Hairetu(I - 1 + TateMin, J - 1 + YokoMin) 'Hairetuの中の値
        Next J
    Next I
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示するときに各列の幅を同じに整えるために
    '文字列長さとその各列の最大値を計算する。
    '以下では「Hairetu」は扱わず、「WithTableHairetu」を扱う。
    N = UBound(WithTableHairetu, 1) '「WithTableHairetu」の縦インデックス数（行数）
    M = UBound(WithTableHairetu, 2) '「WithTableHairetu」の横インデックス数（列数）
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr As String
    For J = 1 To M
        For I = 1 To N
        
            If J > 1 And HyoujiMaxNagasa <> 0 Then
                '最大表示長さが指定されている場合。
                '1列目のテーブルはそのままにする。
                TmpStr = WithTableHairetu(I, J)
                WithTableHairetu(I, J) = 文字列を指定バイト数文字数に省略(TmpStr, HyoujiMaxNagasa)
            End If
            
            NagasaList(I, J) = LenB(StrConv(WithTableHairetu(I, J), vbFromUnicode)) '全角と半角を区別して長さを計算する。
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示するために" "(半角スペース)を追加して
    '文字列長さを同じにする。
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa As Long
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) 'その列の最大文字列長さ
        For I = 1 To N
            'Rept…指定文字列を指定個数連続してつなげた文字列を出力する。
            '（最大文字数-文字数）の分" "（半角スペース）を後ろにくっつける。
            NagasaOnajiList(I, J) = WithTableHairetu(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示する文字列を作成
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
    
    ''※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示
    Debug.Print HairetuName
    For I = 1 To N
        Debug.Print OutputList(I)
    Next I
    
End Sub

Private Function 文字列を指定バイト数文字数に省略(Mojiretu As String, ByteNum As Integer)
    '20201023追加
    '文字列を指定省略バイト文字数までの長さで省略する。
    '省略された文字列の最後の文字は"."に変更する。
    '例：Mojiretu = "魑魅魍魎" , ByteNum = 6 … 出力 = "魑魅.."
    '例：Mojiretu = "魑魅魍魎" , ByteNum = 7 … 出力 = "魑魅魍."
    '例：Mojiretu = "魑魅XX魎" , ByteNum = 6 … 出力 = "魑魅X."
    '例：Mojiretu = "魑魅XX魎" , ByteNum = 7 … 出力 = "魑魅XX."
    
    Dim OriginByte As Integer '入力した文字列「Mojiretu」のバイト文字数
    Dim Output                '出力する変数を格納
    
    '「Mojiretu」のバイト文字数計算
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    
    If OriginByte <= ByteNum Then
        '「Mojiretu」のバイト文字数計算が省略するバイト文字数以下なら
        '省略はしない
        Output = Mojiretu
    Else
    
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = 文字列の各文字累計バイト数計算(Mojiretu)
        BunkaiMojiretu = 文字列分解(Mojiretu)
        
        Dim AddMoji As String
        AddMoji = "."
        
        Dim I As Long, N As Long
        N = Len(Mojiretu)
        
        For I = 1 To N
            If RuikeiByteList(I) < ByteNum Then
                Output = Output & BunkaiMojiretu(I)
                
            ElseIf RuikeiByteList(I) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(I), vbFromUnicode)) = 1 Then
                    '例：Mojiretu = "魑魅魍魎" , ByteNum = 6 ,RuikeiByteList(3) = 6
                    'Output = "魑魅.."
                    Output = Output & AddMoji
                Else
                    '例：Mojiretu = "魑魅XX魎" , ByteNum = 6 ,RuikeiByteList(4) = 6
                    'Output = "魑魅X."
                    Output = Output & AddMoji & AddMoji
                End If
                
                Exit For
                
            ElseIf RuikeiByteList(I) > ByteNum Then
                '例：Mojiretu = "魑魅魍魎" , ByteNum = 7 ,RuikeiByteList(4) = 8
                'Output = "魑魅魍."
                Output = Output & AddMoji
                Exit For
            End If
        Next I
        
    End If
        
    文字列を指定バイト数文字数に省略 = Output

    
End Function

Private Function 文字列の各文字累計バイト数計算(Mojiretu As String)
    '20201023追加

    '文字列を1文字ずつに分解して、各文字のバイト文字長を計算し、
    'その累計値を計算する。
    '例：Mojiretu="新型EKワゴン"
    '出力→Output = (2,4,5,6,7,10,12)
    
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
    
    文字列の各文字累計バイト数計算 = Output
    
End Function

Private Function 文字列分解(Mojiretu As String)
    '20201023追加

    '文字列を1文字ずつ分解して配列に格納
    Dim I     As Long
    Dim N     As Long
    Dim Output
    
    N = Len(Mojiretu)
    ReDim Output(1 To N)
    For I = 1 To N
        Output(I) = Mid(Mojiretu, I, 1)
    Next I
    
    文字列分解 = Output
    
End Function

Private Function ExtractArray1D(Array1D, StartNum As Long, EndNum As Long)
'一次元配列の指定範囲を配列として抽出する
'20211009

'引数
'Array1D ・・・一次元配列
'StartNum・・・抽出範囲の開始番号
'EndNum  ・・・抽出範囲の終了番号
                                   
    '引数チェック
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I As Long
    Dim N As Long
    N = UBound(Array1D, 1) '要素数
    
    If StartNum > EndNum Then
        MsgBox ("抽出範囲の開始位置「StartNum」は、終了位置「EndNum」以下でなければなりません")
        Stop
        Exit Function
    ElseIf StartNum < 1 Then
        MsgBox ("抽出範囲の開始位置「StartNum」は1以上の値を入れてください")
        Stop
        Exit Function
    ElseIf EndNum > N Then
        MsgBox ("抽出範囲の終了行「EndNum」は抽出元の一次元配列の要素数" & N & "以下の値を入れてください")
        Stop
        Exit Function
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndNum - StartNum + 1)
    
    For I = StartNum To EndNum
        Output(I - StartNum + 1) = Array1D(I)
    Next I
    
    '出力
    ExtractArray1D = Output
    
End Function

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function ExtractArray(Array2D, StartRow As Long, StartCol As Long, EndRow As Long, EndCol As Long)
'二次元配列の指定範囲を配列として抽出する
'20210917

'引数
'Array2D ・・・二次元配列
'StartRow・・・抽出範囲の開始行番号
'StartCol・・・抽出範囲の開始列番号
'EndRow  ・・・抽出範囲の終了行番号
'EndCol  ・・・抽出範囲の終了列番号
                                   
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
    
    If StartRow > EndRow Then
        MsgBox ("抽出範囲の開始行「StartRow」は、終了行「EndRow」以下でなければなりません")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("抽出範囲の開始列「StartCol」は、終了列「EndCol」以下でなければなりません")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("抽出範囲の開始行「StartRow」は1以上の値を入れてください")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("抽出範囲の開始列「StartCol」は1以上の値を入れてください")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("抽出範囲の終了行「StartRow」は抽出元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("抽出範囲の終了列「StartCol」は抽出元の二次元配列の列数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '出力
    ExtractArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function UnionArray1D(UpperArray1D, LowerArray1D)
'一次元配列同士を結合して1つの配列とする。
'20210923

'UpperArray1D・・・上に結合する一次元配列
'LowerArray1D・・・下に結合する一次元配列

    '引数チェック
    Call CheckArray1D(UpperArray1D, "UpperArray1D")
    Call CheckArray1DStart1(UpperArray1D, "UpperArray1D")
    Call CheckArray1D(LowerArray1D, "LowerArray1D")
    Call CheckArray1DStart1(LowerArray1D, "LowerArray1D")
    
    '処理
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
    
    '出力
    UnionArray1D = Output
    
End Function

Private Function F_Minverse(ByVal Matrix)
    '20210603改良
    'F_Minverse(input_M)
    'F_Minverse(配列)
    '余因子行列を用いて逆行列を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Matrix)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I        As Integer
    Dim J        As Integer
    Dim N        As Integer
    Dim Output() As Double
    N = UBound(Matrix, 1)
    ReDim Output(1 To N, 1 To N)
    
    Dim detM As Double '行列式の値を格納
    detM = F_MDeterm(Matrix) '行列式を求める
    
    Dim Mjyokyo '指定の列・行を除去した配列を格納
    
    For I = 1 To N '各列
        For J = 1 To N '各行
            
            'I列,J行を除去する
            Mjyokyo = F_Mjyokyo(Matrix, J, I)
            
            'I列,J行の余因子を求めて出力する逆行列に格納
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_Minverse = Output
    
End Function

Private Sub 正方行列かチェック(Matrix)
    '20210603追加
    
    If UBound(Matrix, 1) <> UBound(Matrix, 2) Then
        MsgBox ("正方行列を入力してください" & vbLf & _
                "入力された配列の要素数は" & "「" & _
                UBound(Matrix, 1) & "×" & UBound(Matrix, 2) & "」" & "です")
        Stop
        End
    End If

End Sub

Private Function F_MDeterm(Matrix)
    '20210603改良
    'F_MDeterm(Matrix)
    'F_MDeterm(配列)
    '行列式を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Matrix)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim N As Integer
    N = UBound(Matrix, 1)
    
    Dim Matrix2 '掃き出しを行う行列
    Matrix2 = Matrix
    
    For I = 1 To N '各列
        For J = I To N '掃き出し元の行の探索
            If Matrix2(J, I) <> 0 Then
                K = J '掃き出し元の行
                Exit For
            End If
            
            If J = N And Matrix2(J, I) = 0 Then '掃き出し元の値が全て0なら行列式の値は0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I列,I行)以外で掃き出しとなる場合は行を入れ替え
            Matrix2 = F_Mgyoirekae(Matrix2, I, K)
        End If
        
        '掃き出し
        Matrix2 = F_Mgyohakidasi(Matrix2, I, I)
              
    Next I
    
    
    '行列式の計算
    Dim Output As Double
    Output = 1
    
    For I = 1 To N '各(I列,I行)を掛け合わせていく
        Output = Output * Matrix2(I, I)
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MDeterm = Output
    
End Function

Private Function F_Mgyoirekae(Matrix, Row1 As Integer, Row2 As Integer)
    '20210603改良
    'F_Mgyoirekae(Matrix, Row1, Row2)
    'F_Mgyoirekae(配列,指定行番号①,指定行番号②)
    '行列Matrixの①行と②行を入れ替える
    
    Dim I     As Integer
    Dim J     As Integer
    Dim K     As Integer
    Dim M     As Integer
    Dim N     As Integer
    Dim Output
    
    Output = Matrix
    M = UBound(Matrix, 2) '列数取得
    
    For I = 1 To M
        Output(Row2, I) = Matrix(Row1, I)
        Output(Row1, I) = Matrix(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Private Function F_Mgyohakidasi(Matrix, Row As Integer, Col As Integer)
    '20210603改良
    'F_Mgyohakidasi(Matrix, Row, Col)
    'F_Mgyohakidasi(配列,指定行,指定列)
    '行列MatrixのRow行､Col列の値で各行を掃き出す
    
    Dim I     As Integer
    Dim J     As Integer
    Dim N     As Integer
    Dim Output
    
    Output = Matrix
    N = UBound(Output, 1) '行数取得
    
    Dim Hakidasi '掃き出し元の行
    Dim X As Double '掃き出し元の値
    Dim Y As Double
    ReDim Hakidasi(1 To N)
    X = Matrix(Row, Col)
    
    For I = 1 To N '掃き出し元の1行を作成
        Hakidasi(I) = Matrix(Row, I)
    Next I
    
    For I = 1 To N '各行
        If I = Row Then
            '掃き出し元の行の場合はそのまま
            For J = 1 To N
                Output(I, J) = Matrix(I, J)
            Next J
        
        Else
            '掃き出し元の行以外の場合は掃き出し
            Y = Matrix(I, Col) '掃き出し基準の列の値
            For J = 1 To N
                Output(I, J) = Matrix(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Private Function F_Mjyokyo(Matrix, Row As Integer, Col As Integer)
    '20210603改良
    'F_Mjyokyo(Matrix, Row, Col)
    'F_Mjyokyo(配列,指定行,指定列)
    '行列MatrixのRow行、Col列を除去した行列を返す
    
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim M As Integer
    Dim N As Integer '数え上げ用(Integer型)
    Dim Output '指定した行・列を除去後の配列
    
    N = UBound(Matrix, 1) '行数取得
    M = UBound(Matrix, 2) '列数取得
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2 As Integer
    Dim J2 As Integer
    
    I2 = 0 '行方向数え上げ初期化
    For I = 1 To N
        If I = Row Then
            'なにもしない
        Else
            I2 = I2 + 1 '行方向数え上げ
            
            J2 = 0 '列方向数え上げ初期化
            For J = 1 To M
                If J = Col Then
                    'なにもしない
                Else
                    J2 = J2 + 1 '列方向数え上げ
                    Output(I2, J2) = Matrix(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function

Private Function F_MMult(ByVal Matrix1, ByVal Matrix2)
    'F_MMult(Matrix1, Matrix2)
    'F_MMult(配列①,配列②)
    '行列の積を計算
    '20180213改良
    '20210603改良
    
    '入力値のチェックと修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '配列の次元チェック
    Dim JigenCheck1 As Integer
    Dim JigenCheck2 As Integer
    On Error Resume Next
    JigenCheck1 = UBound(Matrix1, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(Matrix2, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が1なら次元2にする。例)配列(1 to N)→配列(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Matrix1 = Application.Transpose(Matrix1)
    End If
    If IsEmpty(JigenCheck2) Then
        Matrix2 = Application.Transpose(Matrix2)
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If UBound(Matrix1, 1) = 0 Or UBound(Matrix1, 2) = 0 Then
        Matrix1 = Application.Transpose(Application.Transpose(Matrix1))
    End If
    If UBound(Matrix2, 1) = 0 Or UBound(Matrix2, 2) = 0 Then
        Matrix2 = Application.Transpose(Application.Transpose(Matrix2))
    End If
    
    '入力値のチェック
    If UBound(Matrix1, 2) <> UBound(Matrix2, 1) Then
        MsgBox ("配列1の列数と配列2の行数が一致しません。" & vbLf & _
               "(出力) = (配列1)(配列2)")
        Stop
        End
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I        As Integer
    Dim J        As Integer
    Dim K        As Integer
    Dim M        As Integer
    Dim N        As Integer
    Dim M2       As Integer
    Dim Output() As Double '出力する配列
    N = UBound(Matrix1, 1) '配列1の行数
    M = UBound(Matrix1, 2) '配列1の列数
    M2 = UBound(Matrix2, 2) '配列2の列数
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '各行
        For J = 1 To M2 '各列
            For K = 1 To M '(配列1のI行)と(配列2のJ列)を掛け合わせる
                Output(I, J) = Output(I, J) + Matrix1(I, K) * Matrix2(K, J)
            Next K
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MMult = Output
    
End Function


