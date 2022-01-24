<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Print_DataTable
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.RadioButton_PDF = New System.Windows.Forms.RadioButton()
        Me.RadioButton_DOC = New System.Windows.Forms.RadioButton()
        Me.CheckBox_PrintDirectly = New System.Windows.Forms.CheckBox()
        Me.Btn_Print = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ComboBox_InstalledPrinters = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox_PrinterSetting = New System.Windows.Forms.GroupBox()
        Me.TextBox_SelectedPrinter = New System.Windows.Forms.TextBox()
        Me.CheckBox_UserHasPrinter = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel_Setting = New System.Windows.Forms.Panel()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox_PrinterSetting.SuspendLayout()
        Me.Panel_Setting.SuspendLayout()
        Me.SuspendLayout()
        '
        'RadioButton_PDF
        '
        Me.RadioButton_PDF.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton_PDF.AutoSize = True
        Me.RadioButton_PDF.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.RadioButton_PDF.Location = New System.Drawing.Point(290, 42)
        Me.RadioButton_PDF.Name = "RadioButton_PDF"
        Me.RadioButton_PDF.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_PDF.Size = New System.Drawing.Size(176, 23)
        Me.RadioButton_PDF.TabIndex = 7
        Me.RadioButton_PDF.Text = "طباعة على شكل ملف PDF"
        Me.RadioButton_PDF.UseVisualStyleBackColor = True
        '
        'RadioButton_DOC
        '
        Me.RadioButton_DOC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton_DOC.AutoSize = True
        Me.RadioButton_DOC.Checked = True
        Me.RadioButton_DOC.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.RadioButton_DOC.Location = New System.Drawing.Point(257, 13)
        Me.RadioButton_DOC.Name = "RadioButton_DOC"
        Me.RadioButton_DOC.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_DOC.Size = New System.Drawing.Size(209, 23)
        Me.RadioButton_DOC.TabIndex = 8
        Me.RadioButton_DOC.TabStop = True
        Me.RadioButton_DOC.Text = "طباعة على شكل ملف وورد DOC"
        Me.RadioButton_DOC.UseVisualStyleBackColor = True
        '
        'CheckBox_PrintDirectly
        '
        Me.CheckBox_PrintDirectly.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBox_PrintDirectly.AutoSize = True
        Me.CheckBox_PrintDirectly.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_PrintDirectly.Location = New System.Drawing.Point(185, 71)
        Me.CheckBox_PrintDirectly.Name = "CheckBox_PrintDirectly"
        Me.CheckBox_PrintDirectly.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox_PrintDirectly.Size = New System.Drawing.Size(281, 23)
        Me.CheckBox_PrintDirectly.TabIndex = 6
        Me.CheckBox_PrintDirectly.Text = "الطباعة مباشرتاً بدون عرض الوثيقة قبل الطباعة"
        Me.CheckBox_PrintDirectly.UseVisualStyleBackColor = True
        '
        'Btn_Print
        '
        Me.Btn_Print.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Btn_Print.Font = New System.Drawing.Font("Tahoma", 12.0!)
        Me.Btn_Print.Image = Global.MS_Word_Print_Test.My.Resources.Resources.Print32
        Me.Btn_Print.Location = New System.Drawing.Point(12, 281)
        Me.Btn_Print.Name = "Btn_Print"
        Me.Btn_Print.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Btn_Print.Size = New System.Drawing.Size(454, 53)
        Me.Btn_Print.TabIndex = 5
        Me.Btn_Print.Text = "Print"
        Me.Btn_Print.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_Print.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.Btn_Print.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 340)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DataGridView1.Size = New System.Drawing.Size(454, 355)
        Me.DataGridView1.TabIndex = 4
        '
        'ComboBox_InstalledPrinters
        '
        Me.ComboBox_InstalledPrinters.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox_InstalledPrinters.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_InstalledPrinters.FormattingEnabled = True
        Me.ComboBox_InstalledPrinters.Location = New System.Drawing.Point(17, 71)
        Me.ComboBox_InstalledPrinters.Name = "ComboBox_InstalledPrinters"
        Me.ComboBox_InstalledPrinters.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ComboBox_InstalledPrinters.Size = New System.Drawing.Size(419, 21)
        Me.ComboBox_InstalledPrinters.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(316, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label1.Size = New System.Drawing.Size(123, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "الطابعة المحددة للطباعة :"
        '
        'GroupBox_PrinterSetting
        '
        Me.GroupBox_PrinterSetting.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox_PrinterSetting.Controls.Add(Me.Panel_Setting)
        Me.GroupBox_PrinterSetting.Controls.Add(Me.Label3)
        Me.GroupBox_PrinterSetting.Controls.Add(Me.CheckBox_UserHasPrinter)
        Me.GroupBox_PrinterSetting.Enabled = False
        Me.GroupBox_PrinterSetting.Location = New System.Drawing.Point(12, 100)
        Me.GroupBox_PrinterSetting.Name = "GroupBox_PrinterSetting"
        Me.GroupBox_PrinterSetting.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.GroupBox_PrinterSetting.Size = New System.Drawing.Size(454, 175)
        Me.GroupBox_PrinterSetting.TabIndex = 11
        Me.GroupBox_PrinterSetting.TabStop = False
        Me.GroupBox_PrinterSetting.Text = "الطابعة المحددة :"
        '
        'TextBox_SelectedPrinter
        '
        Me.TextBox_SelectedPrinter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox_SelectedPrinter.Location = New System.Drawing.Point(17, 31)
        Me.TextBox_SelectedPrinter.Name = "TextBox_SelectedPrinter"
        Me.TextBox_SelectedPrinter.Size = New System.Drawing.Size(419, 20)
        Me.TextBox_SelectedPrinter.TabIndex = 12
        '
        'CheckBox_UserHasPrinter
        '
        Me.CheckBox_UserHasPrinter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBox_UserHasPrinter.Location = New System.Drawing.Point(6, 19)
        Me.CheckBox_UserHasPrinter.Name = "CheckBox_UserHasPrinter"
        Me.CheckBox_UserHasPrinter.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox_UserHasPrinter.Size = New System.Drawing.Size(436, 37)
        Me.CheckBox_UserHasPrinter.TabIndex = 11
        Me.CheckBox_UserHasPrinter.Text = " حدد هذا الخيار للأختيار الطابعة وستتم الطباعة مباشرتاً على الطابعة المحددة" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "أو ل" &
    "ا تحدد إذا كنت تريد إظهار شاشة إعدادات الطباعة قبل الطباعة"
        Me.CheckBox_UserHasPrinter.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(72, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label2.Size = New System.Drawing.Size(367, 26)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "هذا هو اسم الطابعة المحددة للطباعة والذي سيتم إرسالة لأمر الطباعة" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "إذا تركت هذا ا" &
    "لحقل خالي ستتم الطباعة مباشرتاً ولكن على الطابعة الإفتراضية ."
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Label3.Location = New System.Drawing.Point(6, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label3.Size = New System.Drawing.Size(436, 1)
        Me.Label3.TabIndex = 13
        '
        'Panel_Setting
        '
        Me.Panel_Setting.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Setting.Controls.Add(Me.Label2)
        Me.Panel_Setting.Controls.Add(Me.ComboBox_InstalledPrinters)
        Me.Panel_Setting.Controls.Add(Me.TextBox_SelectedPrinter)
        Me.Panel_Setting.Controls.Add(Me.Label1)
        Me.Panel_Setting.Enabled = False
        Me.Panel_Setting.Location = New System.Drawing.Point(6, 63)
        Me.Panel_Setting.Name = "Panel_Setting"
        Me.Panel_Setting.Size = New System.Drawing.Size(442, 106)
        Me.Panel_Setting.TabIndex = 14
        '
        'Frm_Print_DataTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(478, 707)
        Me.Controls.Add(Me.GroupBox_PrinterSetting)
        Me.Controls.Add(Me.RadioButton_PDF)
        Me.Controls.Add(Me.RadioButton_DOC)
        Me.Controls.Add(Me.CheckBox_PrintDirectly)
        Me.Controls.Add(Me.Btn_Print)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Frm_Print_DataTable"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Frm_Print_DataTable"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox_PrinterSetting.ResumeLayout(False)
        Me.Panel_Setting.ResumeLayout(False)
        Me.Panel_Setting.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents RadioButton_PDF As RadioButton
    Friend WithEvents RadioButton_DOC As RadioButton
    Friend WithEvents CheckBox_PrintDirectly As CheckBox
    Friend WithEvents Btn_Print As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ComboBox_InstalledPrinters As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox_PrinterSetting As GroupBox
    Friend WithEvents CheckBox_UserHasPrinter As CheckBox
    Friend WithEvents TextBox_SelectedPrinter As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel_Setting As Panel
    Friend WithEvents Label3 As Label
End Class
