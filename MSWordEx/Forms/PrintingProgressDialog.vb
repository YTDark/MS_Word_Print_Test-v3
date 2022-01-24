Public Class PrintingProgressDialog


    Public Sub SetProgressCuonter()
        Dim NewValue As Double = Me.ProgressBarEx1.Value + 1
        If NewValue > Me.ProgressBarEx1.Maximum Then
            Exit Sub
        End If
        Me.ProgressBarEx1.Value += 1
        Dim Percentage As Double = Me.ProgressBarEx1.Value / Me.ProgressBarEx1.Maximum * 100
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub Btn_CancelRepairing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_CancelRepairing.Click
        MSO.Print.UserRequestCancelPrintingProgress = True
    End Sub

End Class















