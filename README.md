# VBA-Approximate-formula
近似式、補間計算用のVBAモジュール

License: The MIT license

Copyright (c) 2021 YujiFukami

開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

開発テスト環境 OS: Windows 10 Pro

# 使い方
「ModApproximate.bas」「ModImmediate.bas」「ModMatrix.bas」をVBEにインポートすること。

# 使える関数紹介
## Spline(ArrayX1D, ArrayY1D, InputX),SplineXY(ArrayXY2D, InputX)
グラフ（X座標の配列とY座標の配列）の任意X座標でのスプライン補間の値を出力する

関数Splineの各引数

- ArrayX1D ：グラフのX座標の配列（1次元配列）

- ArrayY1D ：グラフのY座標の配列（1次元配列）

- InputX   ：補間位置のX座標


関数SplineXYの各引数

- ArrayXY2D  ：グラフのXY座標の配列（2次元配列）

- InputX   ：補間位置のX座標



関数はワークシート関数としても利用可能

![Spline, SplineXYのグラフ](https://user-images.githubusercontent.com/73621859/128811920-5f08c4ea-b3e9-4140-8d4c-311f2cdf6573.jpg)

![Spline, SplineXYのワークシート関数入力](https://user-images.githubusercontent.com/73621859/128811919-b027ddcc-b751-431d-ba25-8a6fc3240c0f.jpg)


## SplineByArrayX1D(ArrayX1D, ArrayY1D, InputArrayX1D), SplineXYByArrayX1D(ArrayXY2D, InputArrayX1D)
補間位置のX座標を配列で一気に入力して、補間値Yを配列で一気に出力する

関数SplineByArrayX1Dの各引数

- ArrayX1D      ：グラフのX座標の配列（1次元配列）

- ArrayY1D      ：グラフのY座標の配列（1次元配列）

- InputArrayX1D ：補間位置のX座標の配列（1次元配列）


関数SplineXYByArrayX1Dの各引数

- ArrayXY2D     ：グラフのXY座標の配列（2次元配列）

- InputArrayX1D ：補間位置のX座標の配列（1次元配列）


関数はワークシート関数としても利用可能

![SplineByArrayX1D,SplineXYByArrayX1Dのグラフ](https://user-images.githubusercontent.com/73621859/128811939-a2cd2a20-e5af-480b-b384-2b5f03869193.jpg)

![XYグラフ](https://user-images.githubusercontent.com/73621859/128813230-4aecfb81-a978-4c0f-bc6a-67682084f17e.jpg)
![SplineByArrayX1D,SplineXYByArrayXのワークシート関数入力](https://user-images.githubusercontent.com/73621859/128811938-b73fb46f-932e-41ef-8136-509f48b9edda.jpg)

## SplinePara(ArrayX1D, ArrayY1D, BunkatuN), SplineXYPara(ArrayXY2D, BunkatuN)
グラフのX座標の範囲を指定自然数で分割したX座標による補間値Y座標を一気に出力する。

出力値はXY座標の2次元配列

関数SplineParaの各引数

- ArrayX1D ：グラフのX座標の配列（1次元配列）

- ArrayY1D ：グラフのY座標の配列（1次元配列）

- BunkatuN ：X座標範囲を分割する分割数


関数SplineXYParaの各引数

- ArrayXY2D     ：グラフのXY座標の配列（2次元配列）

- BunkatuN ：X座標範囲を分割する分割数


関数はSplineXYParaのみワークシート関数としても利用可能

![SplinePara, SplineXYParaのグラフ](https://user-images.githubusercontent.com/73621859/128811937-61396b6f-712e-4bb6-ad9c-b08a466a7387.jpg)

![XYグラフ](https://user-images.githubusercontent.com/73621859/128813230-4aecfb81-a978-4c0f-bc6a-67682084f17e.jpg)
![SplineXYParaのワークシート関数入力](https://user-images.githubusercontent.com/73621859/128811940-bad25131-eadd-4d73-b3d0-3c492ad9122b.jpg)
