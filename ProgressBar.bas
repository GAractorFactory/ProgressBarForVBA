Attribute VB_Name = "ProgressBar"
'Option Explicit

Sub showProgress()
    frmProgress.Show vbModeless                        ' ���[�h���X�ŃE�B���h�E�i�t�H�[���j��\������
    frmProgress.MousePointer = fmMousePointerHourGlass ' �}�E�X�J�[�\���������v�\���ɂ���
    Call mainProc                                      ' �傽�鏈�����Ăяo��
    frmProgress.MousePointer = fmMousePointerArrow     ' �������I���΃}�E�X�J�[�\������ɖ߂�
    MsgBox "�������I�����܂����B", vbInformation       ' �I�����b�Z�[�W��\������
    Unload frmProgress                                   ' �t�H�[�����B��
End Sub

Sub setupProgressBar(title As String)
    With frmProgress
     .Caption = title               ' �t�H�[���̃^�C�g���̐ݒ�
     .lblStatus.Caption = title     ' �X�e�[�^�X���x���̐ݒ�
     .lblSection.Caption = ""
     .lblPercentage.Caption = "0.0 %"
     .lblCount.Caption = ""
     .BarFg.BackColor = &H8000000D  ' �v���O���X�o�[�̑O�ʐF�̐ݒ�
     .BarFg.Width = 0               ' �v���O���X�o�[�̏����l0��ݒ�
    End With
End Sub

Sub updateStatus(ByVal Filename As String, ByRef num As Integer, ByRef denom As Integer)
On Error GoTo ErrHandler
    If denom = 0 Then GoTo ZeroDivisionException
    If num > denom Then GoTo LargerNumException

' ���v���O���X�o�[�̃��C������
    With frmProgress
      .lblCount.Caption = "(" & num & " / " & denom & ")"          ' �J�E���g��\��
      .lblSection = Filename                                       ' �������̃t�@�C�����̕\��
      .BarFg.Width = .BarBg.Width * (num / denom)                  ' �v���O���X�o�[��i�߂�
      .lblPercentage = Format(((num / denom) * 100), "0.0") & " %" ' �i��������\��
      .Repaint                                                     ' �ĕ`��
    End With

Exit Sub

LargerNumException:
    MsgBox "���q����������傫���ł�", vbCritical
    Exit Sub

ZeroDivisionException:
    MsgBox "�[���ɂ�鏜�Z�ł�", vbCritical
    Exit Sub

ErrHandler:
    MsgBox Err.Number & ": " & Err.Description, vbCritical

End Sub
