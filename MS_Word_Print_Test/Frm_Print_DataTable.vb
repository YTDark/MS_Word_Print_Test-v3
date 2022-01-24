Public Class Frm_Print_DataTable

    Dim Emps_Table As New DataTable

    Private Sub Frm_Print_DataGridView_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Emps_Table.Columns.Add("Emp_Name", GetType(String))
        Emps_Table.Columns.Add("Emp_Age", GetType(String))
        Emps_Table.Columns.Add("Emp_Profession", GetType(String))


        Emps_Table.Rows.Add(New String() {"عبدالله الدوسري", "40", "مبرمج"})
        Emps_Table.Rows.Add(New String() {"خالد سليمان", "46", "محاسب"})
        Emps_Table.Rows.Add(New String() {"سلطان علي", "25", "إداري"})
        Emps_Table.Rows.Add(New String() {"كمال يوسف", "22", "مندوب مبيعات"})
        Emps_Table.Rows.Add(New String() {"احمد محمد", "32", "علاقات عامة"})
        Emps_Table.Rows.Add(New String() {"ياسين مصطفى", "21", "مراسل"})
        Emps_Table.Rows.Add(New String() {"عبدالقادر محمد", "45", "مدير إقليمي"})
        Emps_Table.Rows.Add(New String() {"ناصر عبدالله", "39", "مدير تنفيذي"})
        Emps_Table.Rows.Add(New String() {"سالم البصول", "36", "مشرف عام"})
        Emps_Table.Rows.Add(New String() {"ابو العجب احمد", "51", "مدير عام"})
        Emps_Table.Rows.Add(New String() {"فياض عبدالكريم", "49", "محامي"})
        Emps_Table.Rows.Add(New String() {"ناصر السفياني", "38", "إداري"})

        Me.DataGridView1.DataSource = Emps_Table

        Me.DataGridView1.Columns("Emp_Name").HeaderText = "اسم الموظف"
        Me.DataGridView1.Columns("Emp_Age").HeaderText = "العمر"
        Me.DataGridView1.Columns("Emp_Profession").HeaderText = "المهنة"
        Me.DataGridView1.Columns("Emp_Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridView1.AllowUserToAddRows = False ' لمنع ظهور سجل جديد خالي


        Dim InstalledPrinter As String
        For i = 0 To Printing.PrinterSettings.InstalledPrinters.Count - 1
            InstalledPrinter = Printing.PrinterSettings.InstalledPrinters.Item(i)
            Me.ComboBox_InstalledPrinters.Items.Add(InstalledPrinter)
        Next
        Me.ComboBox_InstalledPrinters.SelectedIndex = 0

    End Sub



    Private Sub Btn_Print_Click(sender As Object, e As EventArgs) Handles Btn_Print.Click


        Dim PrintJob As New MSwordEx.MSO.Printing.PrintJob
        With PrintJob
            .AddText(Now.ToString("dddd dd/MM/yyyyم  hh:mm tt"), "PrintDateTime")
            .AddText("كشف أسماء الموظفين", "ReportTitleName")
            .AddText("إدارة منتدى فيجوال بيسك لكل العرب", "DepartmentName")
            With .AddTable()
                .DataTable = Me.Emps_Table
                '-------------------------------------
                .IsFirstColumnAutoNumber = True
                .TableHeadBookMarkName = "TableHead"
                .FirstRowBookMarkName = "TableFirstRow"
                .DeleteTableIfNoData = False
                '-------------------------------------
                .AddTextColumn("Emp_Name")
                .AddTextColumn("Emp_Age")
                .AddTextColumn("Emp_Profession")
                '-------------------------------------
            End With
        End With

        Dim TemplatePath As String = My.Application.Info.DirectoryPath & "\MSO_Documents\Emps_Report.dotx"
        '--------------------------------------------------------------
        Dim MyTemplate As New MSwordEx.MSO.TemplateInfo(TemplatePath)
        With MyTemplate
            '-------------------------------------
            .Caption = "العنوان الذي سيظهر في أعلى برنامج الوورد إذا ظهر"
            .PrintJob = PrintJob

            ' هذة الخيارات مكتوبة هنا فقط إذا أردت أن تغير فيها
            ' بإمكانك حذها والإستغناء عنها أو التعديل على الخيارات التي تريدها
            ' هذة هي القيم الإفتراضية 
            'With .ViewOptions
            '    '-------------------------------------
            '    .ShowBookmarks = False
            '    .ShowTableGridlines = False
            '    .ArabicNumeral = MSwordEx.MSO.Enums.MSArabicNumeral.NumeralHindi
            '    .DisplayPageBoundaries = True
            '    .NormalViewDisplayRulers = True
            '    .ViewType = MSwordEx.MSO.Enums.MSViewType.PrintPreview
            '    .WindowState = MSwordEx.MSO.Enums.MSWindowState.Maximize
            '    .NormalViewZoomPageFit = MSwordEx.MSO.Enums.MSPageFit.PageFitBestFit
            '    .NormalViewZoomPercentage = Nothing
            '    '-------------------------------------
            '    .PrintPreviewDisplayRulers = True
            '    .PrintPreviewPageFitness = New MSwordEx.MSO.PrintPreviewPageFitness(0, 0)
            '    .PrintPreviewZoomPageFit = MSwordEx.MSO.Enums.MSPageFit.PageFitBestFit
            '    .PrintPreviewZoomPercentage = Nothing
            '    '-------------------------------------
            'End With
        End With

        Module_Print.MSWord.ClearAllDocuments()
        Module_Print.MSWord.TemplatesInfoCollection.Clear()
        Module_Print.MSWord.AddNewTemplateInfo(MyTemplate)
        If Module_Print.MSWord.PrepairTemplatesCollection(False) = False Then
            Exit Sub
        End If

        If Me.RadioButton_PDF.Checked Then
            Dim PDFReportFilename As String = Module_Print.GetNewFileName("Emps_Report-", "pdf")
            Module_Print.MSWord.ExportAsPdfFormat(PDFReportFilename)
            If Me.CheckBox_PrintDirectly.Checked Then
                If Me.CheckBox_UserHasPrinter.Checked Then

                    Module_Print.MSWord.PrintOut(Me.TextBox_SelectedPrinter.Text)

                Else
                    Module_Print.MSWord.PrintShowDialog()
                End If
            Else
                Process.Start(PDFReportFilename)
            End If
        Else ' If DOC
            If Me.CheckBox_PrintDirectly.Checked Then
                If Me.CheckBox_UserHasPrinter.Checked Then
                    Module_Print.MSWord.PrintOut(Me.TextBox_SelectedPrinter.Text)
                Else
                    Module_Print.MSWord.PrintShowDialog()
                End If
            Else
                Module_Print.MSWord.SetMSWordVisible(True)
            End If
        End If

    End Sub

    Private Sub ComboBox_InstalledPrinters_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_InstalledPrinters.SelectedIndexChanged
        Me.TextBox_SelectedPrinter.Text = Me.ComboBox_InstalledPrinters.Text
    End Sub

    Private Sub CheckBox_UserHasPrinter_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_UserHasPrinter.CheckedChanged
        Me.Panel_Setting.Enabled = Me.CheckBox_UserHasPrinter.Checked
    End Sub

    Private Sub CheckBox_PrintDirectly_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_PrintDirectly.CheckedChanged
        Me.GroupBox_PrinterSetting.Enabled = Me.CheckBox_PrintDirectly.Checked
    End Sub
End Class