Attribute VB_Name = "ProgressBar"
'Option Explicit

Sub showProgress()
    frmProgress.Show vbModeless                        ' モードレスでウィンドウ（フォーム）を表示する
    frmProgress.MousePointer = fmMousePointerHourGlass ' マウスカーソルを砂時計表示にする
    Call mainProc                                      ' 主たる処理を呼び出す
    frmProgress.MousePointer = fmMousePointerArrow     ' 処理が終わればマウスカーソルを矢印に戻す
    MsgBox "処理が終了しました。", vbInformation       ' 終了メッセージを表示する
    Unload frmProgress                                   ' フォームを隠す
End Sub

Sub setupProgressBar(title As String)
    With frmProgress
     .Caption = title               ' フォームのタイトルの設定
     .lblStatus.Caption = title     ' ステータスラベルの設定
     .lblSection.Caption = ""
     .lblPercentage.Caption = "0.0 %"
     .lblCount.Caption = ""
     .BarFg.BackColor = &H8000000D  ' プログレスバーの前面色の設定
     .BarFg.Width = 0               ' プログレスバーの初期値0を設定
    End With
End Sub

Sub updateStatus(ByVal Filename As String, ByRef num As Integer, ByRef denom As Integer)
On Error GoTo ErrHandler
    If denom = 0 Then GoTo ZeroDivisionException
    If num > denom Then GoTo LargerNumException

' ★プログレスバーのメイン処理
    With frmProgress
      .lblCount.Caption = "(" & num & " / " & denom & ")"          ' カウントを表示
      .lblSection = Filename                                       ' 処理中のファイル名の表示
      .BarFg.Width = .BarBg.Width * (num / denom)                  ' プログレスバーを進める
      .lblPercentage = Format(((num / denom) * 100), "0.0") & " %" ' 進捗割合を表示
      .Repaint                                                     ' 再描画
    End With

Exit Sub

LargerNumException:
    MsgBox "分子が分母よりも大きいです", vbCritical
    Exit Sub

ZeroDivisionException:
    MsgBox "ゼロによる除算です", vbCritical
    Exit Sub

ErrHandler:
    MsgBox Err.Number & ": " & Err.Description, vbCritical

End Sub
