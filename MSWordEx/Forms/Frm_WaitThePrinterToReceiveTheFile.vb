Imports Microsoft.Office.Interop

Public Class Frm_WaitThePrinterToReceiveTheFile
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.Icon = My.Resources.Printer16
    End Sub

    Dim WordApplication As Word.Application

    Public Function WaitThePrinterToReceiveTheFile(ByRef pWordApplication As Word.Application) As Boolean
        Me.WordApplication = pWordApplication
        Me.ShowDialog()
    End Function

    Private Sub Frm_WaitThePrinterToReceiveTheFile_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Do While Me.WordApplication.BackgroundPrintingStatus > 0
            System.Windows.Forms.Application.DoEvents()
        Loop
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
End Class