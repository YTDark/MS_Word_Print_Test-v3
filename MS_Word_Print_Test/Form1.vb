Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Module_Print.MSWord = New MSwordEx.MSO.MSWord(My.Forms.Form1)
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            If Module_Print.MSWord IsNot Nothing Then
                Module_Print.MSWord.CloseApplication()
            End If
        Catch
        End Try
    End Sub




    Private Sub Btn_Print_DataGridView_Click(sender As Object, e As EventArgs) Handles Btn_Print_DataGridView.Click
        Using Frm As New Frm_Print_DataGridView
            Frm.ShowDialog()
        End Using
    End Sub

    Private Sub Btn_Print_DataTable_Click(sender As Object, e As EventArgs) Handles Btn_Print_DataTable.Click
        Using Frm As New Frm_Print_DataTable
            Frm.ShowDialog()
        End Using
    End Sub

    Private Sub Btn_Print_FixData_Click(sender As Object, e As EventArgs) Handles Btn_Print_FixData.Click
        Using Frm As New Frm_Print_FixData
            Frm.ShowDialog()
        End Using
    End Sub
End Class
