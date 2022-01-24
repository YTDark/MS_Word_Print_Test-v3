Public Class Frm_Print_DataGridView
    Private Sub Frm_Print_DataGridView_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.DataGridView1.AllowUserToAddRows = False ' لمنع ظهور سجل جديد خالي

        Me.DataGridView1.Columns.Add("Emp_Name", "اسم الموظف")
        Me.DataGridView1.Columns.Add("Emp_Age", "العمر")
        Me.DataGridView1.Columns.Add("Emp_Profession", "المهنة")

        Me.DataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        Me.DataGridView1.Rows.Add(New String() {"عبدالله الدوسري", "40", "مبرمج"})
        Me.DataGridView1.Rows.Add(New String() {"خالد سليمان", "46", "محاسب"})
        Me.DataGridView1.Rows.Add(New String() {"سلطان علي", "25", "إداري"})
        Me.DataGridView1.Rows.Add(New String() {"كمال يوسف", "22", "مندوب مبيعات"})
        Me.DataGridView1.Rows.Add(New String() {"احمد محمد", "32", "علاقات عامة"})
        Me.DataGridView1.Rows.Add(New String() {"ياسين مصطفى", "21", "مراسل"})
        Me.DataGridView1.Rows.Add(New String() {"عبدالقادر محمد", "45", "مدير إقليمي"})
        Me.DataGridView1.Rows.Add(New String() {"ناصر عبدالله", "39", "مدير تنفيذي"})
        Me.DataGridView1.Rows.Add(New String() {"سالم البصول", "36", "مشرف عام"})
        Me.DataGridView1.Rows.Add(New String() {"ابو العجب احمد", "51", "مدير عام"})
        Me.DataGridView1.Rows.Add(New String() {"فياض عبدالكريم", "49", "محامي"})
        Me.DataGridView1.Rows.Add(New String() {"ناصر السفياني", "38", "إداري"})

    End Sub

    Private Sub Btn_Print_Click(sender As Object, e As EventArgs) Handles Btn_Print.Click

        Dim PrintJob As New MSwordEx.MSO.Printing.PrintJob
        With PrintJob
            .AddText(Now.ToString("dddd dd/MM/yyyyم  hh:mm tt"), "PrintDateTime")
            .AddText("كشف أسماء الموظفين", "ReportTitleName")
            .AddText("إدارة منتدى فيجوال بيسك لكل العرب", "DepartmentName")
            With .AddTable()
                .DataTable = Me.CraeteReportTable
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
                Module_Print.MSWord.PrintShowDialog()
            Else
                Process.Start(PDFReportFilename)
            End If
        Else ' If DOC
            If Me.CheckBox_PrintDirectly.Checked Then
                Module_Print.MSWord.PrintShowDialog()
            Else
                Module_Print.MSWord.SetMSWordVisible(True)
            End If
        End If

    End Sub



    ' هذة الوظيفة لإنشاء جدول بنفس شكل بيانات الداتا قريد فيو
    ' DataTable النتيجة ستكون جدول من نوع 
    Private Function CraeteReportTable() As DataTable
        Dim T As New DataTable
        T.Columns.Add("Emp_Name", GetType(String))
        T.Columns.Add("Emp_Age", GetType(String))
        T.Columns.Add("Emp_Profession", GetType(String))
        For Each r As DataGridViewRow In Me.DataGridView1.Rows
            Dim NewRow As DataRow = T.NewRow
            NewRow.Item("Emp_Name") = r.Cells("Emp_Name").Value
            NewRow.Item("Emp_Age") = r.Cells("Emp_Age").Value
            NewRow.Item("Emp_Profession") = r.Cells("Emp_Profession").Value
            T.Rows.Add(NewRow)
        Next
        Return T
    End Function

End Class