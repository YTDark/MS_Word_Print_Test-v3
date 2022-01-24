<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Print_DataGridView
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Btn_Print = New System.Windows.Forms.Button()
        Me.CheckBox_PrintDirectly = New System.Windows.Forms.CheckBox()
        Me.RadioButton_DOC = New System.Windows.Forms.RadioButton()
        Me.RadioButton_PDF = New System.Windows.Forms.RadioButton()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 159)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DataGridView1.Size = New System.Drawing.Size(430, 431)
        Me.DataGridView1.TabIndex = 0
        '
        'Btn_Print
        '
        Me.Btn_Print.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Btn_Print.Font = New System.Drawing.Font("Tahoma", 12.0!)
        Me.Btn_Print.Image = Global.MS_Word_Print_Test.My.Resources.Resources.Print32
        Me.Btn_Print.Location = New System.Drawing.Point(12, 100)
        Me.Btn_Print.Name = "Btn_Print"
        Me.Btn_Print.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Btn_Print.Size = New System.Drawing.Size(430, 53)
        Me.Btn_Print.TabIndex = 1
        Me.Btn_Print.Text = "Print"
        Me.Btn_Print.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_Print.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.Btn_Print.UseVisualStyleBackColor = True
        '
        'CheckBox_PrintDirectly
        '
        Me.CheckBox_PrintDirectly.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBox_PrintDirectly.AutoSize = True
        Me.CheckBox_PrintDirectly.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.CheckBox_PrintDirectly.Location = New System.Drawing.Point(161, 71)
        Me.CheckBox_PrintDirectly.Name = "CheckBox_PrintDirectly"
        Me.CheckBox_PrintDirectly.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox_PrintDirectly.Size = New System.Drawing.Size(281, 23)
        Me.CheckBox_PrintDirectly.TabIndex = 2
        Me.CheckBox_PrintDirectly.Text = "الطباعة مباشرتاً بدون عرض الوثيقة قبل الطباعة"
        Me.CheckBox_PrintDirectly.UseVisualStyleBackColor = True
        '
        'RadioButton_DOC
        '
        Me.RadioButton_DOC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton_DOC.AutoSize = True
        Me.RadioButton_DOC.Checked = True
        Me.RadioButton_DOC.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.RadioButton_DOC.Location = New System.Drawing.Point(233, 13)
        Me.RadioButton_DOC.Name = "RadioButton_DOC"
        Me.RadioButton_DOC.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_DOC.Size = New System.Drawing.Size(209, 23)
        Me.RadioButton_DOC.TabIndex = 3
        Me.RadioButton_DOC.TabStop = True
        Me.RadioButton_DOC.Text = "طباعة على شكل ملف وورد DOC"
        Me.RadioButton_DOC.UseVisualStyleBackColor = True
        '
        'RadioButton_PDF
        '
        Me.RadioButton_PDF.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton_PDF.AutoSize = True
        Me.RadioButton_PDF.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold)
        Me.RadioButton_PDF.Location = New System.Drawing.Point(266, 42)
        Me.RadioButton_PDF.Name = "RadioButton_PDF"
        Me.RadioButton_PDF.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton_PDF.Size = New System.Drawing.Size(176, 23)
        Me.RadioButton_PDF.TabIndex = 3
        Me.RadioButton_PDF.Text = "طباعة على شكل ملف PDF"
        Me.RadioButton_PDF.UseVisualStyleBackColor = True
        '
        'Frm_Print_DataGridView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(454, 602)
        Me.Controls.Add(Me.RadioButton_PDF)
        Me.Controls.Add(Me.RadioButton_DOC)
        Me.Controls.Add(Me.CheckBox_PrintDirectly)
        Me.Controls.Add(Me.Btn_Print)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Frm_Print_DataGridView"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Frm_DataGridView"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Btn_Print As Button
    Friend WithEvents CheckBox_PrintDirectly As CheckBox
    Friend WithEvents RadioButton_DOC As RadioButton
    Friend WithEvents RadioButton_PDF As RadioButton
End Class
