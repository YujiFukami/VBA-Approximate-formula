Attribute VB_Name = "ModApproximate"
'シート関数用近似、補間関数
'2016/12/22更新
'20170622 F_SENKEIHOKANを追加
Function F_SplineXY(ByVal HairetuXY, X#)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力値Xに対する補間値Y
    
    '＜入力値の説明＞
    'HariretuXY：補間の対象となるX,Yの値が格納された配列
    'HairetuXYの1列目がX,2列目がYとなるようにする。
    'X:補間一のXの値
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim Output_Y#
    Output_Y = F_Spline(HairetuX, HairetuY, X)
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_SplineXY = Output_Y
    
End Function

Function F_SplineXYByXList(ByVal HairetuXY, ByVal XList)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力配列XListに対する補間値の配列YList
    
    '＜入力値の説明＞
    'HariretuXY：補間の対象となるX,Yの値が格納された配列
    'HairetuXYの1列目がX,2列目がYとなるようにする。
    'XList:補間位置Xが格納された配列
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim Output_YList
    Output_YList = F_SplineByXList(HairetuX, HairetuY, XList)
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_SplineXYByXList = Output_YList
    
End Function

Function F_SplineXYPara(ByVal HairetuXY, BunkatuN&)
    'パラメトリック関数形式でスプライン補間を行う
    'HairetuX,HairetuYがどちらも単調増加、単調減少でない場合に用いる。
    '＜出力値の説明＞
    'パラメトリック関数形式で保管されたXList,YListが格納されたXYList
    '1列目がXList,2列目がYList
    
    '＜入力値の説明＞
    'HariretuXY：補間の対象となるX,Yの値が格納された配列
    'HairetuXYの1列目がX,2列目がYとなるようにする。
    'パラメトリック関数の分割個数（出力されるXList,YListの要素数は(分割個数+1)）
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    Dim StartNum%
    StartNum = LBound(HairetuXY) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
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
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_SplineXYPara = OutputXYList
    
End Function

Function F_Spline#(ByVal HairetuX, ByVal HairetuY, X#)
        
    '20171124修正
    '20180309改良
    
    'スプライン補間計算を行う
    '
    '<出力値の説明>
    '入力値Xに対する補間値Y
    '
    '<入力値の説明>
    'HairetuX：補間の対象とする配列X
    'HairetuY：補間の対象とする配列Y
    'X：補間位置のXの値
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(HairetuX, 1) <> 1 Then
        HairetuX = Application.Transpose(Application.Transpose(HairetuX))
    End If
    If LBound(HairetuY, 1) <> 1 Then
        HairetuY = Application.Transpose(Application.Transpose(HairetuY))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(HairetuX, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(HairetuY, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        HairetuX = Application.Transpose(HairetuX)
    End If
    If JigenCheck2 > 0 Then
        HairetuY = Application.Transpose(HairetuY)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, N%, K%, A, B, C, D
    Dim Output_Y# '出力値Y
    Dim SotoNaraTrue As Boolean
    SotoNaraTrue = False
    N = UBound(HairetuX, 1)
       
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(HairetuX, HairetuY)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
        
    For I = 1 To N - 1
        If HairetuX(I) < HairetuX(I + 1) Then 'Xが単調増加の場合
            If I = 1 And HairetuX(1) > X Then '範囲に入らないとき(開始点より前)
                Output_Y = HairetuY(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And HairetuX(I + 1) <= X Then '範囲に入らないとき(終了点より後)
                Output_Y = HairetuY(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf HairetuX(I) <= X And HairetuX(I + 1) > X Then '範囲内
                K = I: Exit For
            
            End If
        Else 'Xが単調減少の場合
        
            If I = 1 And HairetuX(1) < X Then '範囲に入らないとき(開始点より前)
                Output_Y = HairetuY(1)
                SotoNaraTrue = True
                Exit For
            
            ElseIf I = N - 1 And HairetuX(I + 1) >= X Then '範囲に入らないとき(終了点より後)
                Output_Y = HairetuY(N)
                SotoNaraTrue = True
                Exit For
                
            ElseIf HairetuX(I + 1) < X And HairetuX(I) >= X Then '範囲内
                K = I: Exit For
            
            End If
        
        End If
    Next I
        
    If SotoNaraTrue = False Then
        Output_Y = A(K) + B(K) * (X - HairetuX(K)) + C(K) * (X - HairetuX(K)) ^ 2 + D(K) * (X - HairetuX(K)) ^ 3
    End If
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_Spline = Output_Y

End Function

Function F_SplinePara(ByVal HairetuX, ByVal HairetuY, BunkatuN&)
    'パラメトリック関数形式でスプライン補間を行う
    'HairetuX,HairetuYがどちらも単調増加、単調減少でない場合に用いる。
    '＜出力値の説明＞
    'パラメトリック関数形式で保管されたXList,YList
    
    '＜入力値の説明＞
    'HariretuX：補間の対象となるXの値が格納された配列
    'HariretuY：補間の対象となるYの値が格納された配列
    'パラメトリック関数の分割個数（出力されるXList,YListの要素数は(分割個数+1)）
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim StartNum%
    '行列の開始要素を1に変更（計算しやすいから）
    StartNum = LBound(HairetuX, 1) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
    If LBound(HairetuX, 1) <> 1 Then
        HairetuX = Application.Transpose(Application.Transpose(HairetuX))
    End If
    If LBound(HairetuY, 1) <> 1 Then
        HairetuY = Application.Transpose(Application.Transpose(HairetuY))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(HairetuX, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(HairetuY, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        HairetuX = Application.Transpose(HairetuX)
    End If
    If JigenCheck2 > 0 Then
        HairetuY = Application.Transpose(HairetuY)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(HairetuX, 1)
    Dim HairetuT#(), TList#()
    
    'X,Yの補間の基準となる配列を作成
    ReDim HairetuT(1 To N)
    For I = 1 To N
        '0～1を等間隔
        HairetuT(I) = (I - 1) / (N - 1)
    Next I
    
    '出力補間位置の基準位置
    If JigenCheck1 > 0 Then '出力値の形状を入力値に合わせるための処理
        ReDim TList(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            TList(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim TList(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            TList(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim Output_XList, Output_YList
    Output_XList = F_SplineByXList(HairetuT, HairetuX, TList)
    Output_YList = F_SplineByXList(HairetuT, HairetuY, TList)
    
    '出力
    Dim Output(1 To 2)
    Output(1) = Output_XList
    Output(2) = Output_YList
    
    F_SplinePara = Output
    
End Function

Function F_SplineByXList(ByVal HairetuX, ByVal HairetuY, ByVal XList)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力配列XListに対する補間値の配列YList
    
    '＜入力値の説明＞
    'HariretuX：補間の対象となるXの値が格納された配列
    'HariretuY：補間の対象となるYの値が格納された配列
    'XList:補間位置Xが格納された配列

    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim StartNum%
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(HairetuX, 1) <> 1 Then
        HairetuX = Application.Transpose(Application.Transpose(HairetuX))
    End If
    If LBound(HairetuY, 1) <> 1 Then
        HairetuY = Application.Transpose(Application.Transpose(HairetuY))
    End If
    StartNum = LBound(XList, 1) 'XListの開始要素番号を取っておく（出力値を合わせるため）
    If LBound(XList, 1) <> 1 Then
        XList = Application.Transpose(Application.Transpose(XList))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(HairetuX, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(HairetuY, 2) '配列の次元が1ならエラーとなる
    JigenCheck3 = UBound(XList, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        HairetuX = Application.Transpose(HairetuX)
    End If
    If JigenCheck2 > 0 Then
        HairetuY = Application.Transpose(HairetuY)
    End If
    If JigenCheck3 > 0 Then
        XList = Application.Transpose(XList)
    End If

    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim N%, K%, A, B, C, D
    
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(HairetuX, HairetuY)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(HairetuX, 1) '補間対象の要素数
    
    Dim Output_YList#() '出力するYListの格納
    Dim NX%
    NX = UBound(XList, 1) '補間位置の個数
    ReDim Output_YList(1 To NX)
    Dim TmpX#, TmpY#
    
    For J = 1 To NX
        TmpX = XList(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If HairetuX(I) < HairetuX(I + 1) Then 'Xが単調増加の場合
                If I = 1 And HairetuX(1) > TmpX Then '範囲に入らないとき(開始点より前)
                    TmpY = HairetuY(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And HairetuX(I + 1) <= TmpX Then '範囲に入らないとき(終了点より後)
                    TmpY = HairetuY(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf HairetuX(I) <= TmpX And HairetuX(I + 1) > TmpX Then '範囲内
                    K = I: Exit For
                
                End If
            Else 'Xが単調減少の場合
            
                If I = 1 And HairetuX(1) < TmpX Then '範囲に入らないとき(開始点より前)
                    TmpY = HairetuY(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And HairetuX(I + 1) >= TmpX Then '範囲に入らないとき(終了点より後)
                    TmpY = HairetuY(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf HairetuX(I + 1) < TmpX And HairetuX(I) >= TmpX Then '範囲内
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - HairetuX(K)) + C(K) * (TmpX - HairetuX(K)) ^ 2 + D(K) * (TmpX - HairetuX(K)) ^ 3
        End If
        
        Output_YList(J) = TmpY
        
    Next J
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim Output
    
    '出力する配列を入力した配列XListの形状に合わせる
    If JigenCheck3 = 1 Then '入力のXListが二次元配列
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

    '参考：http://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim A, B, C, D
    N = UBound(X, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h#()
    Dim Hairetu_L#() '左辺の配列 要素数(1 to N,1 to N)
    Dim Hairetu_R#() '右辺の配列 要素数(1 to N,1 to 1)
    Dim Hairetu_Lm#() '左辺の配列の逆行列 要素数(1 to N,1 to N)
    
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
    
    '右辺の配列の計算
    For I = 1 To N
        If I = 1 Or I = N Then
            Hairetu_R(I, 1) = 0
        Else
            Hairetu_R(I, 1) = 3 * (Y(I + 1) - Y(I)) / h(I) - 3 * (Y(I) - Y(I - 1)) / h(I - 1)
        End If
    Next I
    
    '左辺の配列の計算
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
    
    '左辺の配列の逆行列
    Hairetu_Lm = F_Minverse(Hairetu_L)
    
    'Cの配列を求める
    C = F_MMult(Hairetu_Lm, Hairetu_R)
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
