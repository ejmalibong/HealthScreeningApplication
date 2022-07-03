<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScreenEntry
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
        Me.txtEmployeeScanId = New System.Windows.Forms.TextBox()
        Me.txtEmployeeCode = New System.Windows.Forms.Label()
        Me.lblEmployeeCode = New System.Windows.Forms.Label()
        Me.txtDate = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.lblEmployeeName = New System.Windows.Forms.Label()
        Me.lblReason = New System.Windows.Forms.Label()
        Me.lblEmployeeScanId = New System.Windows.Forms.Label()
        Me.txtReason = New System.Windows.Forms.TextBox()
        Me.lblAbsentFrom = New System.Windows.Forms.Label()
        Me.lblAbsentTo = New System.Windows.Forms.Label()
        Me.txtAbsentFrom = New System.Windows.Forms.MaskedTextBox()
        Me.txtAbsentTo = New System.Windows.Forms.MaskedTextBox()
        Me.lblNotFtw = New System.Windows.Forms.Label()
        Me.chkNotFtw = New System.Windows.Forms.CheckBox()
        Me.btnDelete = New PinkieControls.ButtonXP()
        Me.btnClear = New PinkieControls.ButtonXP()
        Me.btnClose = New PinkieControls.ButtonXP()
        Me.btnSave = New PinkieControls.ButtonXP()
        Me.lblDiagnosis = New System.Windows.Forms.Label()
        Me.txtDiagnosis = New System.Windows.Forms.TextBox()
        Me.txtEmployeeName = New System.Windows.Forms.TextBox()
        Me.lblClinicClearance = New System.Windows.Forms.Label()
        Me.lblLeaveType = New System.Windows.Forms.Label()
        Me.cmbLeaveType = New System.Windows.Forms.ComboBox()
        Me.lblTotalDays = New System.Windows.Forms.Label()
        Me.txtTotalDays = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtEmployeeScanId
        '
        Me.txtEmployeeScanId.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmployeeScanId.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtEmployeeScanId.Font = New System.Drawing.Font("Verdana", 30.0!)
        Me.txtEmployeeScanId.Location = New System.Drawing.Point(11, 87)
        Me.txtEmployeeScanId.MaxLength = 8
        Me.txtEmployeeScanId.Name = "txtEmployeeScanId"
        Me.txtEmployeeScanId.ShortcutsEnabled = False
        Me.txtEmployeeScanId.Size = New System.Drawing.Size(439, 56)
        Me.txtEmployeeScanId.TabIndex = 0
        Me.txtEmployeeScanId.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEmployeeCode
        '
        Me.txtEmployeeCode.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.txtEmployeeCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmployeeCode.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.txtEmployeeCode.ForeColor = System.Drawing.Color.Black
        Me.txtEmployeeCode.Location = New System.Drawing.Point(110, 145)
        Me.txtEmployeeCode.Name = "txtEmployeeCode"
        Me.txtEmployeeCode.Size = New System.Drawing.Size(340, 24)
        Me.txtEmployeeCode.TabIndex = 525
        Me.txtEmployeeCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.txtEmployeeCode.UseCompatibleTextRendering = True
        '
        'lblEmployeeCode
        '
        Me.lblEmployeeCode.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmployeeCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmployeeCode.ForeColor = System.Drawing.Color.Black
        Me.lblEmployeeCode.Location = New System.Drawing.Point(11, 145)
        Me.lblEmployeeCode.Name = "lblEmployeeCode"
        Me.lblEmployeeCode.Size = New System.Drawing.Size(100, 24)
        Me.lblEmployeeCode.TabIndex = 524
        Me.lblEmployeeCode.Text = " ID Number"
        Me.lblEmployeeCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDate
        '
        Me.txtDate.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.txtDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDate.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.txtDate.ForeColor = System.Drawing.Color.Black
        Me.txtDate.Location = New System.Drawing.Point(110, 197)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(340, 24)
        Me.txtDate.TabIndex = 523
        Me.txtDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.txtDate.UseCompatibleTextRendering = True
        '
        'lblDate
        '
        Me.lblDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDate.ForeColor = System.Drawing.Color.Black
        Me.lblDate.Location = New System.Drawing.Point(11, 197)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(100, 24)
        Me.lblDate.TabIndex = 521
        Me.lblDate.Text = " Date"
        Me.lblDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEmployeeName
        '
        Me.lblEmployeeName.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmployeeName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmployeeName.ForeColor = System.Drawing.Color.Black
        Me.lblEmployeeName.Location = New System.Drawing.Point(11, 171)
        Me.lblEmployeeName.Name = "lblEmployeeName"
        Me.lblEmployeeName.Size = New System.Drawing.Size(100, 24)
        Me.lblEmployeeName.TabIndex = 520
        Me.lblEmployeeName.Text = " Name"
        Me.lblEmployeeName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblReason
        '
        Me.lblReason.BackColor = System.Drawing.SystemColors.Control
        Me.lblReason.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblReason.ForeColor = System.Drawing.Color.Black
        Me.lblReason.Location = New System.Drawing.Point(11, 275)
        Me.lblReason.Name = "lblReason"
        Me.lblReason.Size = New System.Drawing.Size(439, 24)
        Me.lblReason.TabIndex = 519
        Me.lblReason.Text = " Reason / Chief Complaint"
        Me.lblReason.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEmployeeScanId
        '
        Me.lblEmployeeScanId.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmployeeScanId.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmployeeScanId.ForeColor = System.Drawing.Color.Black
        Me.lblEmployeeScanId.Location = New System.Drawing.Point(11, 64)
        Me.lblEmployeeScanId.Name = "lblEmployeeScanId"
        Me.lblEmployeeScanId.Size = New System.Drawing.Size(439, 24)
        Me.lblEmployeeScanId.TabIndex = 518
        Me.lblEmployeeScanId.Text = " Employee ID"
        Me.lblEmployeeScanId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtReason
        '
        Me.txtReason.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtReason.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.txtReason.Location = New System.Drawing.Point(11, 298)
        Me.txtReason.Multiline = True
        Me.txtReason.Name = "txtReason"
        Me.txtReason.Size = New System.Drawing.Size(439, 50)
        Me.txtReason.TabIndex = 4
        '
        'lblAbsentFrom
        '
        Me.lblAbsentFrom.BackColor = System.Drawing.SystemColors.Control
        Me.lblAbsentFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAbsentFrom.ForeColor = System.Drawing.Color.Black
        Me.lblAbsentFrom.Location = New System.Drawing.Point(11, 249)
        Me.lblAbsentFrom.Name = "lblAbsentFrom"
        Me.lblAbsentFrom.Size = New System.Drawing.Size(100, 24)
        Me.lblAbsentFrom.TabIndex = 528
        Me.lblAbsentFrom.Text = " Absent From"
        Me.lblAbsentFrom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAbsentTo
        '
        Me.lblAbsentTo.BackColor = System.Drawing.SystemColors.Control
        Me.lblAbsentTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAbsentTo.ForeColor = System.Drawing.Color.Black
        Me.lblAbsentTo.Location = New System.Drawing.Point(249, 249)
        Me.lblAbsentTo.Name = "lblAbsentTo"
        Me.lblAbsentTo.Size = New System.Drawing.Size(66, 24)
        Me.lblAbsentTo.TabIndex = 530
        Me.lblAbsentTo.Text = " To"
        Me.lblAbsentTo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAbsentFrom
        '
        Me.txtAbsentFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAbsentFrom.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.txtAbsentFrom.Location = New System.Drawing.Point(110, 249)
        Me.txtAbsentFrom.Mask = "00/00/0000"
        Me.txtAbsentFrom.Name = "txtAbsentFrom"
        Me.txtAbsentFrom.PromptChar = Global.Microsoft.VisualBasic.ChrW(45)
        Me.txtAbsentFrom.Size = New System.Drawing.Size(136, 24)
        Me.txtAbsentFrom.TabIndex = 2
        Me.txtAbsentFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtAbsentFrom.ValidatingType = GetType(Date)
        '
        'txtAbsentTo
        '
        Me.txtAbsentTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAbsentTo.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.txtAbsentTo.Location = New System.Drawing.Point(314, 249)
        Me.txtAbsentTo.Mask = "00/00/0000"
        Me.txtAbsentTo.Name = "txtAbsentTo"
        Me.txtAbsentTo.PromptChar = Global.Microsoft.VisualBasic.ChrW(45)
        Me.txtAbsentTo.RejectInputOnFirstFailure = True
        Me.txtAbsentTo.Size = New System.Drawing.Size(136, 24)
        Me.txtAbsentTo.TabIndex = 3
        Me.txtAbsentTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtAbsentTo.ValidatingType = GetType(Date)
        '
        'lblNotFtw
        '
        Me.lblNotFtw.BackColor = System.Drawing.SystemColors.Control
        Me.lblNotFtw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNotFtw.ForeColor = System.Drawing.Color.Black
        Me.lblNotFtw.Location = New System.Drawing.Point(11, 425)
        Me.lblNotFtw.Name = "lblNotFtw"
        Me.lblNotFtw.Size = New System.Drawing.Size(166, 24)
        Me.lblNotFtw.TabIndex = 531
        Me.lblNotFtw.Text = "     Unfit To Work    (F11)"
        Me.lblNotFtw.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkNotFtw
        '
        Me.chkNotFtw.AutoSize = True
        Me.chkNotFtw.BackColor = System.Drawing.SystemColors.Control
        Me.chkNotFtw.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkNotFtw.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.chkNotFtw.Location = New System.Drawing.Point(17, 430)
        Me.chkNotFtw.Name = "chkNotFtw"
        Me.chkNotFtw.Size = New System.Drawing.Size(15, 14)
        Me.chkNotFtw.TabIndex = 6
        Me.chkNotFtw.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnDelete.DefaultScheme = False
        Me.btnDelete.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnDelete.Hint = "Delete existing record"
        Me.btnDelete.Image = Global.SickLeaveScreening.My.Resources.Resources.Delete_16_x_16
        Me.btnDelete.Location = New System.Drawing.Point(124, 467)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnDelete.Size = New System.Drawing.Size(110, 36)
        Me.btnDelete.TabIndex = 540
        Me.btnDelete.TabStop = False
        Me.btnDelete.Text = "Delete (F4)"
        '
        'btnClear
        '
        Me.btnClear.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnClear.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnClear.DefaultScheme = False
        Me.btnClear.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnClear.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnClear.Hint = "Reset data entry form"
        Me.btnClear.Location = New System.Drawing.Point(237, 467)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnClear.Size = New System.Drawing.Size(105, 36)
        Me.btnClear.TabIndex = 538
        Me.btnClear.TabStop = False
        Me.btnClear.Text = "Clear (ESC)"
        '
        'btnClose
        '
        Me.btnClose.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnClose.DefaultScheme = False
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnClose.Hint = "Close data entry form"
        Me.btnClose.Location = New System.Drawing.Point(345, 467)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnClose.Size = New System.Drawing.Size(105, 36)
        Me.btnClose.TabIndex = 539
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnSave.DefaultScheme = False
        Me.btnSave.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnSave.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnSave.Hint = "Save current record"
        Me.btnSave.Image = Global.SickLeaveScreening.My.Resources.Resources.Save_16_x_16
        Me.btnSave.Location = New System.Drawing.Point(11, 467)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnSave.Size = New System.Drawing.Size(110, 36)
        Me.btnSave.TabIndex = 537
        Me.btnSave.TabStop = False
        Me.btnSave.Text = "Save (F10)"
        '
        'lblDiagnosis
        '
        Me.lblDiagnosis.BackColor = System.Drawing.SystemColors.Control
        Me.lblDiagnosis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDiagnosis.ForeColor = System.Drawing.Color.Black
        Me.lblDiagnosis.Location = New System.Drawing.Point(11, 350)
        Me.lblDiagnosis.Name = "lblDiagnosis"
        Me.lblDiagnosis.Size = New System.Drawing.Size(439, 24)
        Me.lblDiagnosis.TabIndex = 542
        Me.lblDiagnosis.Text = " Diagnosis"
        Me.lblDiagnosis.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDiagnosis
        '
        Me.txtDiagnosis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDiagnosis.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.txtDiagnosis.Location = New System.Drawing.Point(11, 373)
        Me.txtDiagnosis.Multiline = True
        Me.txtDiagnosis.Name = "txtDiagnosis"
        Me.txtDiagnosis.Size = New System.Drawing.Size(439, 50)
        Me.txtDiagnosis.TabIndex = 5
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmployeeName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtEmployeeName.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.txtEmployeeName.Location = New System.Drawing.Point(110, 171)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.Size = New System.Drawing.Size(340, 24)
        Me.txtEmployeeName.TabIndex = 1
        '
        'lblClinicClearance
        '
        Me.lblClinicClearance.AutoSize = True
        Me.lblClinicClearance.Font = New System.Drawing.Font("Verdana", 20.0!, System.Drawing.FontStyle.Bold)
        Me.lblClinicClearance.Location = New System.Drawing.Point(104, 18)
        Me.lblClinicClearance.Name = "lblClinicClearance"
        Me.lblClinicClearance.Size = New System.Drawing.Size(273, 32)
        Me.lblClinicClearance.TabIndex = 544
        Me.lblClinicClearance.Text = "Health Screening"
        '
        'lblLeaveType
        '
        Me.lblLeaveType.BackColor = System.Drawing.SystemColors.Control
        Me.lblLeaveType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLeaveType.ForeColor = System.Drawing.Color.Black
        Me.lblLeaveType.Location = New System.Drawing.Point(11, 223)
        Me.lblLeaveType.Name = "lblLeaveType"
        Me.lblLeaveType.Size = New System.Drawing.Size(100, 24)
        Me.lblLeaveType.TabIndex = 545
        Me.lblLeaveType.Text = " Leave Type"
        Me.lblLeaveType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbLeaveType
        '
        Me.cmbLeaveType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLeaveType.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.cmbLeaveType.FormattingEnabled = True
        Me.cmbLeaveType.Location = New System.Drawing.Point(110, 223)
        Me.cmbLeaveType.Name = "cmbLeaveType"
        Me.cmbLeaveType.Size = New System.Drawing.Size(340, 24)
        Me.cmbLeaveType.TabIndex = 546
        '
        'lblTotalDays
        '
        Me.lblTotalDays.BackColor = System.Drawing.SystemColors.Control
        Me.lblTotalDays.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotalDays.ForeColor = System.Drawing.Color.Black
        Me.lblTotalDays.Location = New System.Drawing.Point(331, 425)
        Me.lblTotalDays.Name = "lblTotalDays"
        Me.lblTotalDays.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblTotalDays.Size = New System.Drawing.Size(60, 24)
        Me.lblTotalDays.TabIndex = 547
        Me.lblTotalDays.Text = "Day(s)"
        Me.lblTotalDays.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTotalDays
        '
        Me.txtTotalDays.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalDays.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTotalDays.Enabled = False
        Me.txtTotalDays.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.txtTotalDays.Location = New System.Drawing.Point(390, 425)
        Me.txtTotalDays.Name = "txtTotalDays"
        Me.txtTotalDays.Size = New System.Drawing.Size(60, 24)
        Me.txtTotalDays.TabIndex = 548
        Me.txtTotalDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'frmScreenEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(462, 511)
        Me.Controls.Add(Me.txtTotalDays)
        Me.Controls.Add(Me.lblTotalDays)
        Me.Controls.Add(Me.lblReason)
        Me.Controls.Add(Me.lblDiagnosis)
        Me.Controls.Add(Me.txtReason)
        Me.Controls.Add(Me.txtDiagnosis)
        Me.Controls.Add(Me.lblLeaveType)
        Me.Controls.Add(Me.cmbLeaveType)
        Me.Controls.Add(Me.lblClinicClearance)
        Me.Controls.Add(Me.txtEmployeeName)
        Me.Controls.Add(Me.chkNotFtw)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblNotFtw)
        Me.Controls.Add(Me.txtAbsentTo)
        Me.Controls.Add(Me.lblAbsentFrom)
        Me.Controls.Add(Me.txtAbsentFrom)
        Me.Controls.Add(Me.lblAbsentTo)
        Me.Controls.Add(Me.txtEmployeeScanId)
        Me.Controls.Add(Me.txtEmployeeCode)
        Me.Controls.Add(Me.lblEmployeeCode)
        Me.Controls.Add(Me.txtDate)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.lblEmployeeName)
        Me.Controls.Add(Me.lblEmployeeScanId)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmScreenEntry"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtEmployeeScanId As System.Windows.Forms.TextBox
    Friend WithEvents txtEmployeeCode As System.Windows.Forms.Label
    Friend WithEvents lblEmployeeCode As System.Windows.Forms.Label
    Friend WithEvents txtDate As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblEmployeeName As System.Windows.Forms.Label
    Friend WithEvents lblReason As System.Windows.Forms.Label
    Friend WithEvents lblEmployeeScanId As System.Windows.Forms.Label
    Friend WithEvents txtReason As System.Windows.Forms.TextBox
    Friend WithEvents lblAbsentFrom As System.Windows.Forms.Label
    Friend WithEvents lblAbsentTo As System.Windows.Forms.Label
    Friend WithEvents txtAbsentFrom As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtAbsentTo As System.Windows.Forms.MaskedTextBox
    Friend WithEvents lblNotFtw As System.Windows.Forms.Label
    Friend WithEvents chkNotFtw As System.Windows.Forms.CheckBox
    Friend WithEvents btnDelete As PinkieControls.ButtonXP
    Friend WithEvents btnClear As PinkieControls.ButtonXP
    Friend WithEvents btnClose As PinkieControls.ButtonXP
    Friend WithEvents btnSave As PinkieControls.ButtonXP
    Friend WithEvents lblDiagnosis As System.Windows.Forms.Label
    Friend WithEvents txtDiagnosis As System.Windows.Forms.TextBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents lblClinicClearance As System.Windows.Forms.Label
    Friend WithEvents lblLeaveType As System.Windows.Forms.Label
    Friend WithEvents cmbLeaveType As System.Windows.Forms.ComboBox
    Friend WithEvents lblTotalDays As Label
    Friend WithEvents txtTotalDays As TextBox
End Class
