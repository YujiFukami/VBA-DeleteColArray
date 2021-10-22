Attribute VB_Name = "ModDeleteColArray"
Option Explicit

'DeleteColArray    ・・・元場所：FukamiAddins3.ModArray
'CheckArray2D      ・・・元場所：FukamiAddins3.ModArray
'CheckArray2DStart1・・・元場所：FukamiAddins3.ModArray



Public Function DeleteColArray(Array2D, DeleteCol As Long)
'二次元配列の指定列を消去した配列を出力する
'20210917

'引数
'Array2D  ・・・二次元配列
'DeleteCol・・・消去する列番号

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数

    If DeleteCol < 1 Then
        MsgBox ("削除する列番号は1以上の値を入れてください")
        Stop
        End
    ElseIf DeleteCol > M Then
        MsgBox ("削除する列番号は元の二次元配列の列数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To N, 1 To M - 1)
    For I = 1 To N
        K = 0
        For J = 1 To M
            If J <> DeleteCol Then
                K = K + 1
                Output(I, K) = Array2D(I, J)
            End If
        Next J
    Next I
    
    '出力
    DeleteColArray = Output

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


