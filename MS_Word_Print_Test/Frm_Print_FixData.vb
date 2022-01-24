Public Class Frm_Print_FixData
    Private Sub Frm_Print_FixData_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim FontFilePath As String = My.Application.Info.DirectoryPath & "\fre3of9x.ttf"
        Dim pfc As New Drawing.Text.PrivateFontCollection()
        pfc.AddFontFile(FontFilePath)
        lbl_Barcode_1.Font = New Font(pfc.Families(0), Me.lbl_Barcode_1.Font.Size, FontStyle.Regular)

        Me.lbl_Barcode_2A.Font = New Font(pfc.Families(0), Me.lbl_Barcode_2A.Font.Size, FontStyle.Regular)
        Me.lbl_Barcode_2B.Font = New Font(pfc.Families(0), Me.lbl_Barcode_2A.Font.Size, FontStyle.Regular)
        Me.lbl_Barcode_2C.Font = New Font(pfc.Families(0), Me.lbl_Barcode_2A.Font.Size, FontStyle.Regular)

    End Sub

    Private Sub Btn_Print_Click(sender As Object, e As EventArgs) Handles Btn_Print.Click


        Dim Barcode_1 As Bitmap
        With Me.lbl_Barcode_1
            Barcode_1 = New Bitmap(.Width, .Height)
            .DrawToBitmap(Barcode_1, New Rectangle(0, 0, .Width, .Height))
        End With

        Dim Barcode_2 As Bitmap
        With Me.Panel_Barcode_2
            Barcode_2 = New Bitmap(.Width, .Height)
            Me.lbl_Barcode_2B.SendToBack()
            .DrawToBitmap(Barcode_2, New Rectangle(0, 0, .Width, .Height))
            Me.lbl_Barcode_2B.BringToFront()
        End With



        Dim PrintJob As New MSwordEx.MSO.Printing.PrintJob
        With PrintJob
            .AddText(Me.TextBox_CompanyName.Text, "CompanyName")
            .AddText(Me.TextBox_Address.Text, "Address")
            .AddText(Me.TextBox_City.Text, "City")
            .AddText(Me.TextBox_OrderNumber.Text, "OrderNumber")
            .AddText(Me.TextBox_DeliveryInstruction.Text, "DeliveryInstruction")
            '-----------------------------------------------------------------------
            .AddText(Me.TextBox_OrderNumber.Text, "BarcodeNumber")
            .AddText(Me.TextBox_OrderNumber.Text, "RefNumber")

            With .AddImage()
                .BookMarkName = "BarcodeImage"
                .Image = Barcode_1
            End With

            With .AddImage()
                .BookMarkName = "BarcodeLargImage"
                .Image = Barcode_2
            End With

        End With

        Dim TemplatePath As String = My.Application.Info.DirectoryPath & "\MSO_Documents\INVOICE.dotx"
        '--------------------------------------------------------------
        Dim MyTemplate As New MSwordEx.MSO.TemplateInfo(TemplatePath)
        With MyTemplate
            .Caption = "العنوان الذي سيظهر في أعلى برنامج الوورد إذا ظهر"
            .PrintJob = PrintJob
        End With

        Module_Print.MSWord.ClearAllDocuments()
        Module_Print.MSWord.TemplatesInfoCollection.Clear()
        Module_Print.MSWord.AddNewTemplateInfo(MyTemplate)
        If Module_Print.MSWord.PrepairTemplatesCollection(False) = False Then
            Exit Sub
        End If

        Dim PDFReportFilename As String = Module_Print.GetNewFileName("Emps_Report-", "pdf")
        Module_Print.MSWord.ExportAsPdfFormat(PDFReportFilename)
        Process.Start(PDFReportFilename)

    End Sub

    Private Sub TextBox_OrderNumber_TextChanged(sender As Object, e As EventArgs) Handles TextBox_OrderNumber.TextChanged

        Dim OrderNo As String = "*" & Me.TextBox_OrderNumber.Text.Trim & "*"
        Me.lbl_Barcode_1.Text = OrderNo
        Me.lbl_Barcode_2A.Text = OrderNo
        Me.lbl_Barcode_2B.Text = OrderNo
        Me.lbl_Barcode_2C.Text = OrderNo
    End Sub
End Class