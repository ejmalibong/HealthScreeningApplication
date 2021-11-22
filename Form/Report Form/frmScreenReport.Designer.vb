<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScreenReport
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
        Me.pnlLeft = New System.Windows.Forms.Panel()
        Me.cmbEmploymentType = New System.Windows.Forms.ComboBox()
        Me.chkNotFtw = New System.Windows.Forms.CheckBox()
        Me.lblEmployment = New System.Windows.Forms.Label()
        Me.btnRemoveFilters = New PinkieControls.ButtonXP()
        Me.btnClose = New PinkieControls.ButtonXP()
        Me.btnGenerate = New PinkieControls.ButtonXP()
        Me.cmbLeaveType = New System.Windows.Forms.ComboBox()
        Me.lblLeaveType = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.lblScreeningDate = New System.Windows.Forms.Label()
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
        Me.rptViewer = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.pnlLeft.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlLeft
        '
        Me.pnlLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlLeft.Controls.Add(Me.cmbEmploymentType)
        Me.pnlLeft.Controls.Add(Me.chkNotFtw)
        Me.pnlLeft.Controls.Add(Me.lblEmployment)
        Me.pnlLeft.Controls.Add(Me.btnRemoveFilters)
        Me.pnlLeft.Controls.Add(Me.btnClose)
        Me.pnlLeft.Controls.Add(Me.btnGenerate)
        Me.pnlLeft.Controls.Add(Me.cmbLeaveType)
        Me.pnlLeft.Controls.Add(Me.lblLeaveType)
        Me.pnlLeft.Controls.Add(Me.lblEndDate)
        Me.pnlLeft.Controls.Add(Me.dtpEndDate)
        Me.pnlLeft.Controls.Add(Me.lblStartDate)
        Me.pnlLeft.Controls.Add(Me.lblScreeningDate)
        Me.pnlLeft.Controls.Add(Me.dtpStartDate)
        Me.pnlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlLeft.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.pnlLeft.Location = New System.Drawing.Point(0, 0)
        Me.pnlLeft.Name = "pnlLeft"
        Me.pnlLeft.Size = New System.Drawing.Size(230, 601)
        Me.pnlLeft.TabIndex = 1
        '
        'cmbEmploymentType
        '
        Me.cmbEmploymentType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEmploymentType.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.cmbEmploymentType.FormattingEnabled = True
        Me.cmbEmploymentType.Location = New System.Drawing.Point(12, 225)
        Me.cmbEmploymentType.Name = "cmbEmploymentType"
        Me.cmbEmploymentType.Size = New System.Drawing.Size(205, 24)
        Me.cmbEmploymentType.TabIndex = 159
        '
        'chkNotFtw
        '
        Me.chkNotFtw.AutoSize = True
        Me.chkNotFtw.Location = New System.Drawing.Point(12, 259)
        Me.chkNotFtw.Name = "chkNotFtw"
        Me.chkNotFtw.Size = New System.Drawing.Size(141, 18)
        Me.chkNotFtw.TabIndex = 158
        Me.chkNotFtw.Text = "Unfit To Work Only"
        Me.chkNotFtw.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkNotFtw.UseVisualStyleBackColor = True
        '
        'lblEmployment
        '
        Me.lblEmployment.AutoSize = True
        Me.lblEmployment.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblEmployment.Location = New System.Drawing.Point(7, 207)
        Me.lblEmployment.Name = "lblEmployment"
        Me.lblEmployment.Size = New System.Drawing.Size(84, 14)
        Me.lblEmployment.TabIndex = 157
        Me.lblEmployment.Text = "Employment"
        '
        'btnRemoveFilters
        '
        Me.btnRemoveFilters.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnRemoveFilters.DefaultScheme = False
        Me.btnRemoveFilters.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnRemoveFilters.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnRemoveFilters.Hint = "Remove Filter"
        Me.btnRemoveFilters.Location = New System.Drawing.Point(12, 344)
        Me.btnRemoveFilters.Name = "btnRemoveFilters"
        Me.btnRemoveFilters.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnRemoveFilters.Size = New System.Drawing.Size(205, 40)
        Me.btnRemoveFilters.TabIndex = 155
        Me.btnRemoveFilters.TabStop = False
        Me.btnRemoveFilters.Text = "Clear Filter"
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnClose.DefaultScheme = False
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnClose.Hint = "Close"
        Me.btnClose.Location = New System.Drawing.Point(12, 389)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnClose.Size = New System.Drawing.Size(205, 40)
        Me.btnClose.TabIndex = 154
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        '
        'btnGenerate
        '
        Me.btnGenerate.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnGenerate.DefaultScheme = False
        Me.btnGenerate.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnGenerate.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnGenerate.Hint = "Generate Report"
        Me.btnGenerate.Location = New System.Drawing.Point(12, 299)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnGenerate.Size = New System.Drawing.Size(205, 40)
        Me.btnGenerate.TabIndex = 153
        Me.btnGenerate.TabStop = False
        Me.btnGenerate.Text = "Generate"
        '
        'cmbLeaveType
        '
        Me.cmbLeaveType.DisplayMember = "LeaveTypeName"
        Me.cmbLeaveType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLeaveType.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.cmbLeaveType.FormattingEnabled = True
        Me.cmbLeaveType.Location = New System.Drawing.Point(12, 171)
        Me.cmbLeaveType.Name = "cmbLeaveType"
        Me.cmbLeaveType.Size = New System.Drawing.Size(205, 24)
        Me.cmbLeaveType.TabIndex = 2
        '
        'lblLeaveType
        '
        Me.lblLeaveType.AutoSize = True
        Me.lblLeaveType.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblLeaveType.Location = New System.Drawing.Point(6, 153)
        Me.lblLeaveType.Name = "lblLeaveType"
        Me.lblLeaveType.Size = New System.Drawing.Size(78, 14)
        Me.lblLeaveType.TabIndex = 26
        Me.lblLeaveType.Text = "Leave Type"
        '
        'lblEndDate
        '
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblEndDate.Location = New System.Drawing.Point(26, 101)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(21, 14)
        Me.lblEndDate.TabIndex = 24
        Me.lblEndDate.Text = "To"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = "MMMM dd, yyyy"
        Me.dtpEndDate.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndDate.Location = New System.Drawing.Point(12, 118)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(205, 22)
        Me.dtpEndDate.TabIndex = 1
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblStartDate.Location = New System.Drawing.Point(26, 54)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(38, 14)
        Me.lblStartDate.TabIndex = 21
        Me.lblStartDate.Text = "From"
        '
        'lblScreeningDate
        '
        Me.lblScreeningDate.AutoSize = True
        Me.lblScreeningDate.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblScreeningDate.Location = New System.Drawing.Point(6, 35)
        Me.lblScreeningDate.Name = "lblScreeningDate"
        Me.lblScreeningDate.Size = New System.Drawing.Size(103, 14)
        Me.lblScreeningDate.TabIndex = 18
        Me.lblScreeningDate.Text = "Screening Date"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = "MMMM dd, yyyy"
        Me.dtpStartDate.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartDate.Location = New System.Drawing.Point(12, 71)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(205, 22)
        Me.dtpStartDate.TabIndex = 0
        '
        'rptViewer
        '
        Me.rptViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rptViewer.Location = New System.Drawing.Point(230, 0)
        Me.rptViewer.Name = "rptViewer"
        Me.rptViewer.Size = New System.Drawing.Size(1084, 601)
        Me.rptViewer.TabIndex = 2
        Me.rptViewer.TabStop = False
        '
        'frmScreenReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(1314, 601)
        Me.Controls.Add(Me.rptViewer)
        Me.Controls.Add(Me.pnlLeft)
        Me.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmScreenReport"
        Me.ShowIcon = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Monitoring Report"
        Me.pnlLeft.ResumeLayout(False)
        Me.pnlLeft.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlLeft As System.Windows.Forms.Panel
    Friend WithEvents btnRemoveFilters As PinkieControls.ButtonXP
    Friend WithEvents btnClose As PinkieControls.ButtonXP
    Friend WithEvents btnGenerate As PinkieControls.ButtonXP
    Friend WithEvents cmbLeaveType As System.Windows.Forms.ComboBox
    Friend WithEvents lblLeaveType As System.Windows.Forms.Label
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents lblScreeningDate As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents rptViewer As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents lblEmployment As System.Windows.Forms.Label
    Friend WithEvents chkNotFtw As System.Windows.Forms.CheckBox
    Friend WithEvents cmbEmploymentType As System.Windows.Forms.ComboBox
End Class
