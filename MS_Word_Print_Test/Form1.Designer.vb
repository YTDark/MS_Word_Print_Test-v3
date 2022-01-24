<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Btn_Print_DataGridView = New System.Windows.Forms.Button()
        Me.Btn_Print_DataTable = New System.Windows.Forms.Button()
        Me.Btn_Print_FixData = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Btn_Print_DataGridView
        '
        Me.Btn_Print_DataGridView.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Btn_Print_DataGridView.Location = New System.Drawing.Point(44, 46)
        Me.Btn_Print_DataGridView.Name = "Btn_Print_DataGridView"
        Me.Btn_Print_DataGridView.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Btn_Print_DataGridView.Size = New System.Drawing.Size(251, 100)
        Me.Btn_Print_DataGridView.TabIndex = 0
        Me.Btn_Print_DataGridView.Text = "طباعة DataGrigView"
        Me.Btn_Print_DataGridView.UseVisualStyleBackColor = True
        '
        'Btn_Print_DataTable
        '
        Me.Btn_Print_DataTable.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Btn_Print_DataTable.Location = New System.Drawing.Point(301, 46)
        Me.Btn_Print_DataTable.Name = "Btn_Print_DataTable"
        Me.Btn_Print_DataTable.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Btn_Print_DataTable.Size = New System.Drawing.Size(251, 100)
        Me.Btn_Print_DataTable.TabIndex = 0
        Me.Btn_Print_DataTable.Text = "طباعة DataTable"
        Me.Btn_Print_DataTable.UseVisualStyleBackColor = True
        '
        'Btn_Print_FixData
        '
        Me.Btn_Print_FixData.Font = New System.Drawing.Font("Arial", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Btn_Print_FixData.Location = New System.Drawing.Point(558, 46)
        Me.Btn_Print_FixData.Name = "Btn_Print_FixData"
        Me.Btn_Print_FixData.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Btn_Print_FixData.Size = New System.Drawing.Size(251, 100)
        Me.Btn_Print_FixData.TabIndex = 0
        Me.Btn_Print_FixData.Text = "طباعة بيانات ثابتة"
        Me.Btn_Print_FixData.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(872, 194)
        Me.Controls.Add(Me.Btn_Print_FixData)
        Me.Controls.Add(Me.Btn_Print_DataTable)
        Me.Controls.Add(Me.Btn_Print_DataGridView)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Btn_Print_DataGridView As Button
    Friend WithEvents Btn_Print_DataTable As Button
    Friend WithEvents Btn_Print_FixData As Button
End Class
