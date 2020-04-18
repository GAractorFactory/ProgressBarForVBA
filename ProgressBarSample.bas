Attribute VB_Name = "ProgressBarSample"
Sub start()
    setupProgressBar "ˆ—’†..."
    Call showProgress
End Sub

Sub mainProc()

Dim num As Integer

For i = 0 To 100
    num = i
    updateStatus ("file" & num & ".xlsx"), num, 100
    Application.Wait Now() + TimeValue("00:00:01")
    DoEvents
Next i

End Sub
