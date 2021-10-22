Attribute VB_Name = "ModDeleteColArray"
Option Explicit

'DeleteColArray    �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2D      �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2DStart1�E�E�E���ꏊ�FFukamiAddins3.ModArray



Public Function DeleteColArray(Array2D, DeleteCol As Long)
'�񎟌��z��̎w�������������z����o�͂���
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'DeleteCol�E�E�E���������ԍ�

    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��

    If DeleteCol < 1 Then
        MsgBox ("�폜�����ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf DeleteCol > M Then
        MsgBox ("�폜�����ԍ��͌��̓񎟌��z��̗�" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
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
    
    '�o��
    DeleteColArray = Output

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


