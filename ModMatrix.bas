Attribute VB_Name = "ModMatrix"
Option Explicit
'行列を使った計算
'代替関数
Function 逆行列(Hairetu)
    逆行列 = F_Minverse(Hairetu)
End Function
Function 行列式(Hairetu)
    行列式 = F_MDeterm(Hairetu)
End Function


Function F_MMult(ByVal Hairetu1, ByVal Hairetu2)
    'F_MMult(Hairetu1, Hairetu2)
    'F_MMult(配列①,配列②)
    '行列の積を計算
    '20180213改良
    '20210603改良
    
    '入力値のチェックと修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(Hairetu1, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(Hairetu2, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が1なら次元2にする。例)配列(1 to N)→配列(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Hairetu1 = Application.Transpose(Hairetu1)
    End If
    If IsEmpty(JigenCheck2) Then
        Hairetu2 = Application.Transpose(Hairetu2)
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If UBound(Hairetu1, 1) = 0 Or UBound(Hairetu1, 2) = 0 Then
        Hairetu1 = Application.Transpose(Application.Transpose(Hairetu1))
    End If
    If UBound(Hairetu2, 1) = 0 Or UBound(Hairetu2, 2) = 0 Then
        Hairetu2 = Application.Transpose(Application.Transpose(Hairetu2))
    End If
    
    '入力値のチェック
    If UBound(Hairetu1, 2) <> UBound(Hairetu2, 1) Then
        MsgBox ("配列1の列数と配列2の行数が一致しません。" & vbLf & _
               "(出力) = (配列1)(配列2)")
        Stop
        End
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim M2%
    Dim Output#() '出力する配列
    N = UBound(Hairetu1, 1) '配列1の行数
    M = UBound(Hairetu1, 2) '配列1の列数
    M2 = UBound(Hairetu2, 2) '配列2の列数
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '各行
        For J = 1 To M2 '各列
            For K = 1 To M '(配列1のI行)と(配列2のJ列)を掛け合わせる
                Output(I, J) = Output(I, J) + Hairetu1(I, K) * Hairetu2(K, J)
            Next K
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MMult = Output
    
    End Function
Sub 正方行列かチェック(Hairetu)
    '20210603追加
    
    If UBound(Hairetu, 1) <> UBound(Hairetu, 2) Then
        MsgBox ("正方行列を入力してください" & vbLf & _
                "入力された配列の要素数は" & "「" & _
                UBound(Hairetu, 1) & "×" & UBound(Hairetu, 2) & "」" & "です")
        Stop
        End
    End If

End Sub
Function F_Minverse(ByVal Hairetu)
    '20210603改良
    'F_Minverse(input_M)
    'F_Minverse(配列)
    '余因子行列を用いて逆行列を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Hairetu, 1) <> 1 Or LBound(Hairetu, 2) <> 1 Then
        Hairetu = Application.Transpose(Application.Transpose(Hairetu))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Hairetu)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, M2%, N% '数え上げ用(Integer型)
    N = UBound(Hairetu, 1)
    Dim Output#()
    ReDim Output(1 To N, 1 To N)
    
    Dim detM# '行列式の値を格納
    detM = F_MDeterm(Hairetu) '行列式を求める
    
    Dim Mjyokyo '指定の列・行を除去した配列を格納
    
    For I = 1 To N '各列
        For J = 1 To N '各行
            
            'I列,J行を除去する
            Mjyokyo = F_Mjyokyo(Hairetu, J, I)
            
            'I列,J行の余因子を求めて出力する逆行列に格納
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_Minverse = Output
    
End Function

Function F_MDeterm(Hairetu)
    '20210603改良
    'F_MDeterm(Hairetu)
    'F_MDeterm(配列)
    '行列式を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Hairetu, 1) <> 1 Or LBound(Hairetu, 2) <> 1 Then
        Hairetu = Application.Transpose(Application.Transpose(Hairetu))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Hairetu)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(Hairetu, 1)
    
    Dim Hairetu2 '掃き出しを行う行列
    Hairetu2 = Hairetu
    
    For I = 1 To N '各列
        For J = I To N '掃き出し元の行の探索
            If Hairetu2(J, I) <> 0 Then
                K = J '掃き出し元の行
                Exit For
            End If
            
            If J = N And Hairetu2(J, I) = 0 Then '掃き出し元の値が全て0なら行列式の値は0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I列,I行)以外で掃き出しとなる場合は行を入れ替え
            Hairetu2 = F_Mgyoirekae(Hairetu2, I, K)
        End If
        
        '掃き出し
        Hairetu2 = F_Mgyohakidasi(Hairetu2, I, I)
              
    Next I
    
    
    '行列式の計算
    Dim Output#
    Output = 1
    
    For I = 1 To N '各(I列,I行)を掛け合わせていく
        Output = Output * Hairetu2(I, I)
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MDeterm = Output
    
    End Function


Function F_Mgyoirekae(Hairetu, Row1%, Row2%)
    '20210603改良
    'F_Mgyoirekae(Hairetu, Row1, Row2)
    'F_Mgyoirekae(配列,指定行番号①,指定行番号②)
    '行列Hairetuの①行と②行を入れ替える
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    
    Output = Hairetu
    M = UBound(Hairetu, 2) '列数取得
    
    For I = 1 To M
        Output(Row2, I) = Hairetu(Row1, I)
        Output(Row1, I) = Hairetu(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Function F_Mgyohakidasi(Hairetu, Row%, Col%)
    '20210603改良
    'F_Mgyohakidasi(Hairetu, Row, Col)
    'F_Mgyohakidasi(配列,指定行,指定列)
    '行列HairetuのRow行､Col列の値で各行を掃き出す
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    
    Output = Hairetu
    N = UBound(Output, 1) '行数取得
    
    Dim Hakidasi '掃き出し元の行
    Dim X# '掃き出し元の値
    Dim Y#
    ReDim Hakidasi(1 To N)
    X = Hairetu(Row, Col)
    
    For I = 1 To N '掃き出し元の1行を作成
        Hakidasi(I) = Hairetu(Row, I)
    Next I
    
    
    For I = 1 To N '各行
        If I = Row Then
            '掃き出し元の行の場合はそのまま
            For J = 1 To N
                Output(I, J) = Hairetu(I, J)
            Next J
        
        Else
            '掃き出し元の行以外の場合は掃き出し
            Y = Hairetu(I, Col) '掃き出し基準の列の値
            For J = 1 To N
                Output(I, J) = Hairetu(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Function F_Mjyokyo(Hairetu, Row%, Col%)
    '20210603改良
    'F_Mjyokyo(Hairetu, Row, Col)
    'F_Mjyokyo(配列,指定行,指定列)
    '行列HairetuのRow行、Col列を除去した行列を返す
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output '指定した行・列を除去後の配列
    
    N = UBound(Hairetu, 1) '行数取得
    M = UBound(Hairetu, 2) '列数取得
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2%, J2%
    
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
                    Output(I2, J2) = Hairetu(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function
