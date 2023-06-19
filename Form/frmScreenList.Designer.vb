<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScreenList
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmScreenList))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.bindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.txtTotalPageNumber = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.txtPageNumber = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.tssGo = New System.Windows.Forms.ToolStripSeparator()
        Me.btnGo = New System.Windows.Forms.ToolStripButton()
        Me.txtUsername = New System.Windows.Forms.Label()
        Me.pnlSearchByText = New System.Windows.Forms.Panel()
        Me.txtCommon = New System.Windows.Forms.TextBox()
        Me.btnSearch = New PinkieControls.ButtonXP()
        Me.lblSearchCriteria = New System.Windows.Forms.Label()
        Me.btnReset = New PinkieControls.ButtonXP()
        Me.cmbSearchCriteria = New System.Windows.Forms.ComboBox()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnLogOut = New PinkieControls.ButtonXP()
        Me.btnDoctor = New PinkieControls.ButtonXP()
        Me.btnReport = New PinkieControls.ButtonXP()
        Me.btnRefresh = New PinkieControls.ButtonXP()
        Me.btnClose = New PinkieControls.ButtonXP()
        Me.btnAdd = New PinkieControls.ButtonXP()
        Me.btnDelete = New PinkieControls.ButtonXP()
        Me.btnEdit = New PinkieControls.ButtonXP()
        Me.dgvList = New System.Windows.Forms.DataGridView()
        Me.ColScreenId = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColScreenDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColReason = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColDiagnosis = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColLeaveTypeId = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColAbsentFrom = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColAbsentTo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColQuantity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColLeaveTypeName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColsFitToWork = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.pnlDate = New System.Windows.Forms.Panel()
        Me.dtpAbsentTo = New System.Windows.Forms.DateTimePicker()
        Me.lblAbsentTo = New System.Windows.Forms.Label()
        Me.lblAbsentFrom = New System.Windows.Forms.Label()
        Me.dtpAbsentFrom = New System.Windows.Forms.DateTimePicker()
        Me.pnlSearchByCmb = New System.Windows.Forms.Panel()
        Me.cmbCommon = New System.Windows.Forms.ComboBox()
        Me.pnlTop.SuspendLayout()
        CType(Me.bindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.bindingNavigator.SuspendLayout()
        Me.pnlSearchByText.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlDate.SuspendLayout()
        Me.pnlSearchByCmb.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTop.Controls.Add(Me.bindingNavigator)
        Me.pnlTop.Controls.Add(Me.txtUsername)
        Me.pnlTop.Controls.Add(Me.pnlSearchByText)
        Me.pnlTop.Controls.Add(Me.btnSearch)
        Me.pnlTop.Controls.Add(Me.lblSearchCriteria)
        Me.pnlTop.Controls.Add(Me.btnReset)
        Me.pnlTop.Controls.Add(Me.cmbSearchCriteria)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1300, 34)
        Me.pnlTop.TabIndex = 0
        '
        'bindingNavigator
        '
        Me.bindingNavigator.AddNewItem = Nothing
        Me.bindingNavigator.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.bindingNavigator.BackColor = System.Drawing.Color.White
        Me.bindingNavigator.CountItem = Me.txtTotalPageNumber
        Me.bindingNavigator.CountItemFormat = "of "
        Me.bindingNavigator.DeleteItem = Nothing
        Me.bindingNavigator.Dock = System.Windows.Forms.DockStyle.None
        Me.bindingNavigator.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.bindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorMovePreviousItem, Me.txtPageNumber, Me.txtTotalPageNumber, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.tssGo, Me.btnGo})
        Me.bindingNavigator.Location = New System.Drawing.Point(1092, 4)
        Me.bindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.bindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.bindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.bindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.bindingNavigator.Name = "bindingNavigator"
        Me.bindingNavigator.PositionItem = Me.txtPageNumber
        Me.bindingNavigator.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.bindingNavigator.Size = New System.Drawing.Size(201, 25)
        Me.bindingNavigator.TabIndex = 10
        Me.bindingNavigator.Text = "PagerPanel"
        '
        'txtTotalPageNumber
        '
        Me.txtTotalPageNumber.Name = "txtTotalPageNumber"
        Me.txtTotalPageNumber.Size = New System.Drawing.Size(21, 22)
        Me.txtTotalPageNumber.Text = "of "
        Me.txtTotalPageNumber.ToolTipText = "Total page number"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'txtPageNumber
        '
        Me.txtPageNumber.AccessibleName = "Position"
        Me.txtPageNumber.AutoSize = False
        Me.txtPageNumber.Name = "txtPageNumber"
        Me.txtPageNumber.Size = New System.Drawing.Size(30, 23)
        Me.txtPageNumber.Text = "0"
        Me.txtPageNumber.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPageNumber.ToolTipText = "Current page"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'tssGo
        '
        Me.tssGo.Name = "tssGo"
        Me.tssGo.Size = New System.Drawing.Size(6, 25)
        '
        'btnGo
        '
        Me.btnGo.AutoSize = False
        Me.btnGo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnGo.Image = CType(resources.GetObject("btnGo.Image"), System.Drawing.Image)
        Me.btnGo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(35, 22)
        Me.btnGo.Text = "Go"
        Me.btnGo.ToolTipText = "Go to page number specified"
        '
        'txtUsername
        '
        Me.txtUsername.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUsername.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.txtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUsername.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.txtUsername.ForeColor = System.Drawing.Color.Black
        Me.txtUsername.Location = New System.Drawing.Point(771, 3)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Padding = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.txtUsername.Size = New System.Drawing.Size(250, 26)
        Me.txtUsername.TabIndex = 527
        Me.txtUsername.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.txtUsername.UseCompatibleTextRendering = True
        '
        'pnlSearchByText
        '
        Me.pnlSearchByText.BackColor = System.Drawing.Color.White
        Me.pnlSearchByText.Controls.Add(Me.txtCommon)
        Me.pnlSearchByText.Location = New System.Drawing.Point(257, 0)
        Me.pnlSearchByText.Name = "pnlSearchByText"
        Me.pnlSearchByText.Size = New System.Drawing.Size(336, 32)
        Me.pnlSearchByText.TabIndex = 166
        Me.pnlSearchByText.Visible = False
        '
        'txtCommon
        '
        Me.txtCommon.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.txtCommon.Location = New System.Drawing.Point(6, 4)
        Me.txtCommon.Name = "txtCommon"
        Me.txtCommon.Size = New System.Drawing.Size(323, 23)
        Me.txtCommon.TabIndex = 0
        '
        'btnSearch
        '
        Me.btnSearch.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnSearch.DefaultScheme = False
        Me.btnSearch.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnSearch.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnSearch.Hint = "Search"
        Me.btnSearch.Image = Global.SickLeaveScreening.My.Resources.Resources.Find_16_x_16
        Me.btnSearch.Location = New System.Drawing.Point(595, 2)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnSearch.Size = New System.Drawing.Size(85, 28)
        Me.btnSearch.TabIndex = 161
        Me.btnSearch.TabStop = False
        Me.btnSearch.Text = "Search"
        '
        'lblSearchCriteria
        '
        Me.lblSearchCriteria.BackColor = System.Drawing.SystemColors.Control
        Me.lblSearchCriteria.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSearchCriteria.ForeColor = System.Drawing.Color.Black
        Me.lblSearchCriteria.Location = New System.Drawing.Point(5, 4)
        Me.lblSearchCriteria.Name = "lblSearchCriteria"
        Me.lblSearchCriteria.Size = New System.Drawing.Size(65, 24)
        Me.lblSearchCriteria.TabIndex = 525
        Me.lblSearchCriteria.Text = "Criteria"
        Me.lblSearchCriteria.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblSearchCriteria.UseCompatibleTextRendering = True
        '
        'btnReset
        '
        Me.btnReset.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnReset.DefaultScheme = False
        Me.btnReset.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnReset.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.btnReset.Hint = "Remove search filter"
        Me.btnReset.Image = Global.SickLeaveScreening.My.Resources.Resources.Undo_16_x_16
        Me.btnReset.Location = New System.Drawing.Point(683, 2)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnReset.Size = New System.Drawing.Size(85, 28)
        Me.btnReset.TabIndex = 162
        Me.btnReset.TabStop = False
        Me.btnReset.Text = "Reset"
        '
        'cmbSearchCriteria
        '
        Me.cmbSearchCriteria.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSearchCriteria.Font = New System.Drawing.Font("Verdana", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSearchCriteria.FormattingEnabled = True
        Me.cmbSearchCriteria.Location = New System.Drawing.Point(69, 4)
        Me.cmbSearchCriteria.Name = "cmbSearchCriteria"
        Me.cmbSearchCriteria.Size = New System.Drawing.Size(185, 24)
        Me.cmbSearchCriteria.TabIndex = 526
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.Color.Gainsboro
        Me.pnlBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlBottom.Controls.Add(Me.lblStatus)
        Me.pnlBottom.Controls.Add(Me.btnLogOut)
        Me.pnlBottom.Controls.Add(Me.btnDoctor)
        Me.pnlBottom.Controls.Add(Me.btnReport)
        Me.pnlBottom.Controls.Add(Me.btnRefresh)
        Me.pnlBottom.Controls.Add(Me.btnClose)
        Me.pnlBottom.Controls.Add(Me.btnAdd)
        Me.pnlBottom.Controls.Add(Me.btnDelete)
        Me.pnlBottom.Controls.Add(Me.btnEdit)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 658)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(1300, 42)
        Me.pnlBottom.TabIndex = 1
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 12.0!)
        Me.lblStatus.ForeColor = System.Drawing.Color.Red
        Me.lblStatus.Location = New System.Drawing.Point(774, 14)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 21)
        Me.lblStatus.TabIndex = 164
        '
        'btnLogOut
        '
        Me.btnLogOut.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnLogOut.DefaultScheme = False
        Me.btnLogOut.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnLogOut.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnLogOut.Hint = "Log out the current session"
        Me.btnLogOut.Location = New System.Drawing.Point(287, 4)
        Me.btnLogOut.Name = "btnLogOut"
        Me.btnLogOut.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnLogOut.Size = New System.Drawing.Size(90, 32)
        Me.btnLogOut.TabIndex = 163
        Me.btnLogOut.TabStop = False
        Me.btnLogOut.Text = "Log Out"
        '
        'btnDoctor
        '
        Me.btnDoctor.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnDoctor.DefaultScheme = False
        Me.btnDoctor.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnDoctor.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnDoctor.Hint = "Doctor masterlist"
        Me.btnDoctor.Image = Global.SickLeaveScreening.My.Resources.Resources.People_16_x_16
        Me.btnDoctor.Location = New System.Drawing.Point(193, 4)
        Me.btnDoctor.Name = "btnDoctor"
        Me.btnDoctor.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnDoctor.Size = New System.Drawing.Size(90, 32)
        Me.btnDoctor.TabIndex = 162
        Me.btnDoctor.TabStop = False
        Me.btnDoctor.Text = "Doctor"
        '
        'btnReport
        '
        Me.btnReport.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnReport.DefaultScheme = False
        Me.btnReport.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnReport.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnReport.Hint = "Generate the monitoring report"
        Me.btnReport.Image = Global.SickLeaveScreening.My.Resources.Resources.Report_16_x_16
        Me.btnReport.Location = New System.Drawing.Point(99, 4)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnReport.Size = New System.Drawing.Size(90, 32)
        Me.btnReport.TabIndex = 161
        Me.btnReport.TabStop = False
        Me.btnReport.Text = "Report"
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnRefresh.DefaultScheme = False
        Me.btnRefresh.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnRefresh.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnRefresh.Hint = "Refresh"
        Me.btnRefresh.Image = Global.SickLeaveScreening.My.Resources.Resources.Refresh_16_x_16
        Me.btnRefresh.Location = New System.Drawing.Point(5, 4)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnRefresh.Size = New System.Drawing.Size(90, 32)
        Me.btnRefresh.TabIndex = 160
        Me.btnRefresh.TabStop = False
        Me.btnRefresh.Text = "Refresh"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnClose.DefaultScheme = False
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnClose.Hint = "Exit the application"
        Me.btnClose.Location = New System.Drawing.Point(1203, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnClose.Size = New System.Drawing.Size(90, 32)
        Me.btnClose.TabIndex = 159
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnAdd.DefaultScheme = False
        Me.btnAdd.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnAdd.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnAdd.Hint = "Add new record"
        Me.btnAdd.Image = Global.SickLeaveScreening.My.Resources.Resources.Create_16_x_16
        Me.btnAdd.Location = New System.Drawing.Point(921, 4)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnAdd.Size = New System.Drawing.Size(90, 32)
        Me.btnAdd.TabIndex = 156
        Me.btnAdd.TabStop = False
        Me.btnAdd.Text = "  Add"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnDelete.DefaultScheme = False
        Me.btnDelete.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnDelete.Hint = "Delete record"
        Me.btnDelete.Image = Global.SickLeaveScreening.My.Resources.Resources.Erase_16_x_16
        Me.btnDelete.Location = New System.Drawing.Point(1109, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnDelete.Size = New System.Drawing.Size(90, 32)
        Me.btnDelete.TabIndex = 158
        Me.btnDelete.TabStop = False
        Me.btnDelete.Text = "Delete"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer), CType(CType(220, Byte), Integer))
        Me.btnEdit.DefaultScheme = False
        Me.btnEdit.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnEdit.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnEdit.Hint = "Modify selected record"
        Me.btnEdit.Image = Global.SickLeaveScreening.My.Resources.Resources.Modify_16_x_16
        Me.btnEdit.Location = New System.Drawing.Point(1015, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnEdit.Size = New System.Drawing.Size(90, 32)
        Me.btnEdit.TabIndex = 157
        Me.btnEdit.TabStop = False
        Me.btnEdit.Text = "  Edit"
        '
        'dgvList
        '
        Me.dgvList.AllowUserToAddRows = False
        Me.dgvList.AllowUserToDeleteRows = False
        Me.dgvList.AllowUserToResizeRows = False
        Me.dgvList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvList.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Verdana", 8.5!)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        Me.dgvList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvList.ColumnHeadersHeight = 25
        Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColScreenId, Me.ColScreenDate, Me.ColEmployeeName, Me.ColReason, Me.ColDiagnosis, Me.ColLeaveTypeId, Me.ColAbsentFrom, Me.ColAbsentTo, Me.ColQuantity, Me.ColLeaveTypeName, Me.ColsFitToWork})
        Me.dgvList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvList.Location = New System.Drawing.Point(0, 34)
        Me.dgvList.MultiSelect = False
        Me.dgvList.Name = "dgvList"
        Me.dgvList.ReadOnly = True
        Me.dgvList.RowHeadersVisible = False
        Me.dgvList.RowHeadersWidth = 40
        Me.dgvList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvList.RowTemplate.DefaultCellStyle.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvList.Size = New System.Drawing.Size(1300, 624)
        Me.dgvList.TabIndex = 2
        '
        'ColScreenId
        '
        Me.ColScreenId.DataPropertyName = "ScreenId"
        Me.ColScreenId.HeaderText = "ScreenId"
        Me.ColScreenId.Name = "ColScreenId"
        Me.ColScreenId.ReadOnly = True
        Me.ColScreenId.Visible = False
        '
        'ColScreenDate
        '
        Me.ColScreenDate.DataPropertyName = "ScreenDate"
        Me.ColScreenDate.HeaderText = "Creation Date"
        Me.ColScreenDate.Name = "ColScreenDate"
        Me.ColScreenDate.ReadOnly = True
        Me.ColScreenDate.Width = 140
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.DataPropertyName = "EmployeeName"
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.Width = 250
        '
        'ColReason
        '
        Me.ColReason.DataPropertyName = "Reason"
        Me.ColReason.HeaderText = "Reason / Chief Complaint"
        Me.ColReason.Name = "ColReason"
        Me.ColReason.ReadOnly = True
        Me.ColReason.Width = 250
        '
        'ColDiagnosis
        '
        Me.ColDiagnosis.DataPropertyName = "Diagnosis"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.ColDiagnosis.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDiagnosis.HeaderText = "Diagnosis"
        Me.ColDiagnosis.Name = "ColDiagnosis"
        Me.ColDiagnosis.ReadOnly = True
        Me.ColDiagnosis.Width = 215
        '
        'ColLeaveTypeId
        '
        Me.ColLeaveTypeId.DataPropertyName = "LeaveTypeId"
        Me.ColLeaveTypeId.HeaderText = "LeaveTypeId"
        Me.ColLeaveTypeId.Name = "ColLeaveTypeId"
        Me.ColLeaveTypeId.ReadOnly = True
        Me.ColLeaveTypeId.Visible = False
        '
        'ColAbsentFrom
        '
        Me.ColAbsentFrom.DataPropertyName = "AbsentFrom"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.Format = "MM/dd/yyyy"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.ColAbsentFrom.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColAbsentFrom.HeaderText = "From"
        Me.ColAbsentFrom.Name = "ColAbsentFrom"
        Me.ColAbsentFrom.ReadOnly = True
        Me.ColAbsentFrom.Width = 110
        '
        'ColAbsentTo
        '
        Me.ColAbsentTo.DataPropertyName = "AbsentTo"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Format = "MM/dd/yyyy"
        Me.ColAbsentTo.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColAbsentTo.HeaderText = "To"
        Me.ColAbsentTo.Name = "ColAbsentTo"
        Me.ColAbsentTo.ReadOnly = True
        Me.ColAbsentTo.Width = 110
        '
        'ColQuantity
        '
        Me.ColQuantity.DataPropertyName = "Quantity"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ColQuantity.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColQuantity.HeaderText = "QTY"
        Me.ColQuantity.Name = "ColQuantity"
        Me.ColQuantity.ReadOnly = True
        Me.ColQuantity.Width = 50
        '
        'ColLeaveTypeName
        '
        Me.ColLeaveTypeName.DataPropertyName = "LeaveTypeCode"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.ColLeaveTypeName.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColLeaveTypeName.HeaderText = "Type"
        Me.ColLeaveTypeName.Name = "ColLeaveTypeName"
        Me.ColLeaveTypeName.ReadOnly = True
        Me.ColLeaveTypeName.Width = 80
        '
        'ColsFitToWork
        '
        Me.ColsFitToWork.DataPropertyName = "IsFitToWork"
        Me.ColsFitToWork.HeaderText = "FTW"
        Me.ColsFitToWork.Name = "ColsFitToWork"
        Me.ColsFitToWork.ReadOnly = True
        Me.ColsFitToWork.Width = 50
        '
        'pnlDate
        '
        Me.pnlDate.BackColor = System.Drawing.Color.White
        Me.pnlDate.Controls.Add(Me.dtpAbsentTo)
        Me.pnlDate.Controls.Add(Me.lblAbsentTo)
        Me.pnlDate.Controls.Add(Me.lblAbsentFrom)
        Me.pnlDate.Controls.Add(Me.dtpAbsentFrom)
        Me.pnlDate.Location = New System.Drawing.Point(258, 1)
        Me.pnlDate.Name = "pnlDate"
        Me.pnlDate.Size = New System.Drawing.Size(336, 32)
        Me.pnlDate.TabIndex = 164
        '
        'dtpAbsentTo
        '
        Me.dtpAbsentTo.CustomFormat = "MMM dd, yyyy"
        Me.dtpAbsentTo.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.dtpAbsentTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAbsentTo.Location = New System.Drawing.Point(200, 5)
        Me.dtpAbsentTo.Name = "dtpAbsentTo"
        Me.dtpAbsentTo.Size = New System.Drawing.Size(130, 22)
        Me.dtpAbsentTo.TabIndex = 21
        '
        'lblAbsentTo
        '
        Me.lblAbsentTo.AutoSize = True
        Me.lblAbsentTo.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.lblAbsentTo.Location = New System.Drawing.Point(177, 9)
        Me.lblAbsentTo.Name = "lblAbsentTo"
        Me.lblAbsentTo.Size = New System.Drawing.Size(21, 14)
        Me.lblAbsentTo.TabIndex = 25
        Me.lblAbsentTo.Text = "To"
        '
        'lblAbsentFrom
        '
        Me.lblAbsentFrom.AutoSize = True
        Me.lblAbsentFrom.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.lblAbsentFrom.Location = New System.Drawing.Point(4, 9)
        Me.lblAbsentFrom.Name = "lblAbsentFrom"
        Me.lblAbsentFrom.Size = New System.Drawing.Size(38, 14)
        Me.lblAbsentFrom.TabIndex = 24
        Me.lblAbsentFrom.Text = "From"
        '
        'dtpAbsentFrom
        '
        Me.dtpAbsentFrom.CustomFormat = "MMM dd, yyyy"
        Me.dtpAbsentFrom.Font = New System.Drawing.Font("Verdana", 9.0!)
        Me.dtpAbsentFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAbsentFrom.Location = New System.Drawing.Point(43, 5)
        Me.dtpAbsentFrom.Name = "dtpAbsentFrom"
        Me.dtpAbsentFrom.Size = New System.Drawing.Size(130, 22)
        Me.dtpAbsentFrom.TabIndex = 20
        '
        'pnlSearchByCmb
        '
        Me.pnlSearchByCmb.BackColor = System.Drawing.Color.White
        Me.pnlSearchByCmb.Controls.Add(Me.cmbCommon)
        Me.pnlSearchByCmb.Location = New System.Drawing.Point(258, 1)
        Me.pnlSearchByCmb.Name = "pnlSearchByCmb"
        Me.pnlSearchByCmb.Size = New System.Drawing.Size(336, 32)
        Me.pnlSearchByCmb.TabIndex = 167
        Me.pnlSearchByCmb.Visible = False
        '
        'cmbCommon
        '
        Me.cmbCommon.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCommon.Font = New System.Drawing.Font("Verdana", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCommon.FormattingEnabled = True
        Me.cmbCommon.Location = New System.Drawing.Point(6, 4)
        Me.cmbCommon.Name = "cmbCommon"
        Me.cmbCommon.Size = New System.Drawing.Size(324, 24)
        Me.cmbCommon.TabIndex = 527
        '
        'frmScreenList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1300, 700)
        Me.Controls.Add(Me.pnlSearchByCmb)
        Me.Controls.Add(Me.pnlDate)
        Me.Controls.Add(Me.dgvList)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlBottom)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(1300, 700)
        Me.Name = "frmScreenList"
        Me.Text = "Health Screening List"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        CType(Me.bindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.bindingNavigator.ResumeLayout(False)
        Me.bindingNavigator.PerformLayout()
        Me.pnlSearchByText.ResumeLayout(False)
        Me.pnlSearchByText.PerformLayout()
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlBottom.PerformLayout()
        CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlDate.ResumeLayout(False)
        Me.pnlDate.PerformLayout()
        Me.pnlSearchByCmb.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents dgvList As System.Windows.Forms.DataGridView
    Friend WithEvents bindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents txtTotalPageNumber As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents txtPageNumber As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents tssGo As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnGo As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnClose As PinkieControls.ButtonXP
    Friend WithEvents btnDelete As PinkieControls.ButtonXP
    Friend WithEvents btnEdit As PinkieControls.ButtonXP
    Friend WithEvents btnAdd As PinkieControls.ButtonXP
    Friend WithEvents btnRefresh As PinkieControls.ButtonXP
    Friend WithEvents btnSearch As PinkieControls.ButtonXP
    Friend WithEvents btnReset As PinkieControls.ButtonXP
    Friend WithEvents lblSearchCriteria As System.Windows.Forms.Label
    Friend WithEvents cmbSearchCriteria As System.Windows.Forms.ComboBox
    Friend WithEvents pnlDate As System.Windows.Forms.Panel
    Friend WithEvents dtpAbsentTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblAbsentTo As System.Windows.Forms.Label
    Friend WithEvents lblAbsentFrom As System.Windows.Forms.Label
    Friend WithEvents dtpAbsentFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents pnlSearchByText As System.Windows.Forms.Panel
    Friend WithEvents txtCommon As System.Windows.Forms.TextBox
    Friend WithEvents pnlSearchByCmb As System.Windows.Forms.Panel
    Friend WithEvents txtUsername As System.Windows.Forms.Label
    Friend WithEvents btnDoctor As PinkieControls.ButtonXP
    Friend WithEvents btnReport As PinkieControls.ButtonXP
    Friend WithEvents btnLogOut As PinkieControls.ButtonXP
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents ColScreenId As DataGridViewTextBoxColumn
    Friend WithEvents ColScreenDate As DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As DataGridViewTextBoxColumn
    Friend WithEvents ColReason As DataGridViewTextBoxColumn
    Friend WithEvents ColDiagnosis As DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveTypeId As DataGridViewTextBoxColumn
    Friend WithEvents ColAbsentFrom As DataGridViewTextBoxColumn
    Friend WithEvents ColAbsentTo As DataGridViewTextBoxColumn
    Friend WithEvents ColQuantity As DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveTypeName As DataGridViewTextBoxColumn
    Friend WithEvents ColsFitToWork As DataGridViewCheckBoxColumn
    Friend WithEvents cmbCommon As ComboBox
End Class
