

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintingProgressDialog
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ProgressBarEx1 = New MSwordEx.ProgressBarEx()
        Me.Tbox_ActionName = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Btn_CancelRepairing = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(270, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label1.Size = New System.Drawing.Size(97, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "جار إعداد التقرير ..."
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.ProgressBarEx1)
        Me.Panel1.Controls.Add(Me.Tbox_ActionName)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(636, 225)
        Me.Panel1.TabIndex = 35
        '
        'ProgressBarEx1
        '
        Me.ProgressBarEx1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBarEx1.BackColor = System.Drawing.Color.FromArgb(CType(CType(215, Byte), Integer), CType(CType(235, Byte), Integer), CType(CType(188, Byte), Integer))
        Me.ProgressBarEx1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ProgressBarEx1.Location = New System.Drawing.Point(35, 126)
        Me.ProgressBarEx1.Maximum = 100
        Me.ProgressBarEx1.Minimum = 0
        Me.ProgressBarEx1.Name = "ProgressBarEx1"
        Me.ProgressBarEx1.ProgressBarColor = System.Drawing.Color.FromArgb(CType(CType(122, Byte), Integer), CType(CType(206, Byte), Integer), CType(CType(66, Byte), Integer))
        Me.ProgressBarEx1.Size = New System.Drawing.Size(566, 22)
        Me.ProgressBarEx1.TabIndex = 38
        Me.ProgressBarEx1.Value = 0
        '
        'Tbox_ActionName
        '
        Me.Tbox_ActionName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Tbox_ActionName.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbox_ActionName.Location = New System.Drawing.Point(105, 100)
        Me.Tbox_ActionName.Name = "Tbox_ActionName"
        Me.Tbox_ActionName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Tbox_ActionName.Size = New System.Drawing.Size(499, 20)
        Me.Tbox_ActionName.TabIndex = 36
        Me.Tbox_ActionName.Text = "عنوان ملف الطباعة"
        Me.Tbox_ActionName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.PictureBox1.Image = Global.MSwordEx.My.Resources.Resources.WindowsLoader
        Me.PictureBox1.Location = New System.Drawing.Point(298, 42)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 36
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Image = Global.MSwordEx.My.Resources.Resources.MicrosoftWordIcon64
        Me.PictureBox2.Location = New System.Drawing.Point(35, 56)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(64, 64)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox2.TabIndex = 5
        Me.PictureBox2.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Btn_CancelRepairing)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 177)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(636, 48)
        Me.Panel2.TabIndex = 35
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Gainsboro
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label3.Location = New System.Drawing.Point(0, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(636, 2)
        Me.Label3.TabIndex = 1
        '
        'Btn_CancelRepairing
        '
        Me.Btn_CancelRepairing.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Btn_CancelRepairing.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Btn_CancelRepairing.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_CancelRepairing.Location = New System.Drawing.Point(513, 9)
        Me.Btn_CancelRepairing.Name = "Btn_CancelRepairing"
        Me.Btn_CancelRepairing.Size = New System.Drawing.Size(88, 30)
        Me.Btn_CancelRepairing.TabIndex = 4
        Me.Btn_CancelRepairing.Text = "إلغاء الأمر"
        Me.Btn_CancelRepairing.UseVisualStyleBackColor = True
        '
        'PrintingProgressDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DodgerBlue
        Me.CancelButton = Me.Btn_CancelRepairing
        Me.ClientSize = New System.Drawing.Size(638, 227)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PrintingProgressDialog"
        Me.Padding = New System.Windows.Forms.Padding(1)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "  "
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Btn_CancelRepairing As System.Windows.Forms.Button
    Friend WithEvents Tbox_ActionName As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ProgressBarEx1 As ProgressBarEx
End Class
