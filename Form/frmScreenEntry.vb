Imports System.Data.SqlClient
Imports BlackCoffeeLibrary
Imports LeaveFilingSystem
Imports LeaveFilingSystem.dsLeaveFiling
Imports LeaveFilingSystem.dsLeaveFilingTableAdapters

Public Class frmScreenEntry
    Private WithEvents absentFrom As Binding
    Private WithEvents absentTo As Binding
    Private WithEvents screenDate As Binding
    Private adpLeaveFiling As New LeaveFilingTableAdapter
    Private adpScreening As New ScreeningTableAdapter
    Private arrSplitted() As String
    Private bsLeaveFiling As New BindingSource
    Private bsScreening As New BindingSource
    Private connection As New clsConnection
    Private dbJeonsoft As New SqlDbMethod(connection.JeonsoftConnection)
    Private dbLeaveFiling As New SqlDbMethod(connection.LocalConnection)
    Private dbMain As New Main

    Private departmentId As Integer = 0
    Private departmentName As String = String.Empty
    Private dsLeaveFiling As New dsLeaveFiling
    Private dtLeaveFiling As New LeaveFilingDataTable
    Private dtScreening As New ScreeningDataTable
    Private employeeId As Integer = 0

    'value from scanner
    Private isDebug As Boolean = SickLeaveScreening.My.MySettings.Default.IsDebug

    Private lstLeaveTypeId As New List(Of Integer)
    Private positionId As Integer = 0
    Private positionName As String = String.Empty
    Private screenBy As Integer = 0
    Private screenId As Integer = 0
    Private teamId As Integer = 0
    Private teamName As String = String.Empty

    'doctor, nurse
    Public Sub New(_screenBy As Integer, Optional _screenId As Integer = 0)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        screenBy = _screenBy
        screenId = _screenId
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ResetForm()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If Not screenId = 0 Then
                Dim rowScreening As ScreeningRow = Me.dsLeaveFiling.Screening.FindByScreenId(screenId)
                Dim count As Integer = 0
                Dim leaveFileId As Integer = 0

                Dim prmCount(0) As SqlParameter
                prmCount(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                prmCount(0).Value = screenId

                count = dbLeaveFiling.ExecuteScalar("SELECT Count(LeaveFileId) FROM dbo.LeaveFiling WHERE ScreenId = @ScreenId", CommandType.Text, prmCount)

                If count > 0 Then
                    MessageBox.Show("Record was already used in the Leave Application System.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                Else
                    If MessageBox.Show("Delete this record?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        Me.bsScreening.RemoveCurrent()
                    End If
                End If

                If Me.dsLeaveFiling.HasChanges Then
                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                    Me.dsLeaveFiling.AcceptChanges()
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                End If
            Else
                If String.IsNullOrEmpty(txtEmployeeCode.Text.Trim) Then
                    Me.ActiveControl = txtEmployeeScanId
                Else
                    Return
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            If String.IsNullOrEmpty(txtEmployeeScanId.Text.Trim) AndAlso String.IsNullOrEmpty(txtEmployeeCode.Text.Trim) Then
                Me.ActiveControl = txtEmployeeScanId
                MessageBox.Show("Please enter employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If cmbLeaveType.SelectedValue = 0 Then
                MessageBox.Show("Please select a leave type.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            If String.IsNullOrEmpty(txtReason.Text.Trim) Then
                MessageBox.Show("Please indicate the reason.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtReason
                Return
            End If

            If String.IsNullOrEmpty(txtDiagnosis.Text.Trim) Then
                MessageBox.Show("Please indicate the diagnosis.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtDiagnosis
                Return
            End If

            If CDate(txtAbsentFrom.Text).Date > CDate(txtAbsentTo.Text).Date Then
                MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtAbsentFrom
                Return
            End If

            'half day leaves
            If (cmbLeaveType.SelectedValue = 12 Or cmbLeaveType.SelectedIndex = 15 Or cmbLeaveType.SelectedValue = 16) AndAlso
                Not (CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date)) Then
                MessageBox.Show("Half-day leave should have the same dates.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            SaveRecord(chkNotFtw.Checked)
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbLeaveType_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbLeaveType.SelectedValueChanged
        If cmbLeaveType.SelectedValue = 14 Then
            chkNotFtw.Enabled = False
            chkNotFtw.CheckState = CheckState.Checked
        Else
            chkNotFtw.Enabled = True
            chkNotFtw.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub dateBinding_Format(sender As Object, e As ConvertEventArgs) Handles screenDate.Format
        If Not e.Value Is DBNull.Value Then
            e.Value = Format(e.Value, "MMMM dd, yyyy  HH:mm")
        Else
            e.Value = dbLeaveFiling.GetServerDate.ToString("MMMM dd, yyyy  HH:mm")
        End If
    End Sub

    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                e.Handled = True
                btnClear.PerformClick()

                screenId = 0
                employeeId = 0
                departmentId = 0
                teamId = 0
                positionId = 0

                txtEmployeeCode.Text = ""
                txtEmployeeName.Clear()
                txtDate.Text = ""
                txtAbsentFrom.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
                txtAbsentFrom.ValidatingType = GetType(System.DateTime)
                txtAbsentTo.ReadOnly = True
                txtAbsentTo.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
                txtAbsentTo.ValidatingType = GetType(System.DateTime)
                txtReason.Clear()
            Case Keys.F4
                e.Handled = True
                btnDelete.PerformClick()
            Case Keys.F10
                e.Handled = True
                btnSave.PerformClick()
            Case Keys.F11
                e.Handled = True
                NotFitToWork()
        End Select
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        dbLeaveFiling.FillCmbWithCaption("SELECT LeaveTypeId, TRIM(LeaveTypeName) AS LeaveTypeName " &
                                         "FROM LeaveType WHERE IsClinic = 1 AND IsActive = 1" &
                                         "ORDER BY TRIM(LeaveTypeName) ",
                                         CommandType.Text, "LeaveTypeId", "LeaveTypeName",
                                         cmbLeaveType, "< Select Leave Type >")

        If Not screenId = 0 Then
            Me.adpScreening.FillByScreenId(Me.dsLeaveFiling.Screening, screenId)
            Me.bsScreening.DataSource = Me.dsLeaveFiling
            Me.bsScreening.DataMember = dtScreening.TableName
            Me.bsScreening.Position = Me.bsScreening.Find("ScreenId", screenId)

            Me.Text = "Record No. " & screenId

            txtEmployeeScanId.Enabled = False
            If CType(Me.bsScreening.Current, DataRowView).Item("EmployeeId") Is DBNull.Value Then
                employeeId = 0
            Else
                employeeId = CType(Me.bsScreening.Current, DataRowView).Item("EmployeeId")
            End If

            txtEmployeeCode.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "EmployeeCode"))
            txtEmployeeName.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "EmployeeName"))
            screenDate = New Binding("Text", Me.bsScreening.Current, "ScreenDate")
            txtDate.DataBindings.Add(screenDate)
            absentFrom = New Binding("Text", Me.bsScreening.Current, "AbsentFrom")
            txtAbsentFrom.DataBindings.Add(absentFrom)
            absentTo = New Binding("Text", Me.bsScreening.Current, "AbsentTo")
            txtAbsentTo.DataBindings.Add(absentTo)
            txtReason.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "Reason"))
            txtDiagnosis.DataBindings.Add(New Binding("Text", Me.bsScreening.Current, "Diagnosis"))

            cmbLeaveType.DataBindings.Add(New Binding("SelectedValue", Me.bsScreening.Current, "LeaveTypeId"))

            If CType(Me.bsScreening.Current, DataRowView).Item("IsFitToWork") = True Then
                chkNotFtw.Checked = False
            Else
                chkNotFtw.Checked = True
            End If

            If txtEmployeeCode.Text.Trim.Substring(0, 3).ToUpper.Trim.Equals("FMB") Then
                txtEmployeeName.ReadOnly = False
            Else
                txtEmployeeName.ReadOnly = True
            End If

            Me.ActiveControl = txtDiagnosis
            txtDiagnosis.Select(txtDiagnosis.Text.Trim.Length, 0)

            txtTotalDays.Text = CType(Me.bsScreening.Current, DataRowView).Item("Quantity")
        Else
            txtTotalDays.Text = 1
            ResetForm()
        End If

        Using reader As IDataReader = dbLeaveFiling.ExecuteReader("SELECT LeaveTypeId FROM dbo.LeaveType WHERE IsClinic = 1 AND LeaveTypeId NOT IN (9,14)",
                                                                  CommandType.Text)
            While reader.Read
                lstLeaveTypeId.Add(reader.Item("LeaveTypeId"))
            End While
            reader.Close()
        End Using
    End Sub

    Private Sub lblNotFtw_Click(sender As Object, e As EventArgs) Handles lblNotFtw.Click
        If chkNotFtw.Enabled = True Then
            If chkNotFtw.CheckState = CheckState.Checked Then
                chkNotFtw.Checked = False
            Else
                chkNotFtw.Checked = True
            End If
        End If
    End Sub

    Private Sub maskedFrom_Format(sender As Object, e As ConvertEventArgs) Handles absentFrom.Format
        e.Value = Format(e.Value, "MM/dd/yyyy")
    End Sub

    Private Sub maskedTo_Format(sender As Object, e As ConvertEventArgs) Handles absentTo.Format
        e.Value = Format(e.Value, "MM/dd/yyyy")
    End Sub

    'validates input from masked textbox - it should be in MM/dd/yyyy format
    Private Sub txtAbsentFrom_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles txtAbsentFrom.TypeValidationCompleted
        If (Not e.IsValidInput) Then
            SendKeys.Send("{End}")
            MessageBox.Show("Please input date in MM/DD/YYYY format.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub txtAbsentFrom_Validated(sender As Object, e As EventArgs) Handles txtAbsentFrom.Validated
        GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
    End Sub

    Private Sub txtAbsentTo_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles txtAbsentTo.TypeValidationCompleted
        If (Not e.IsValidInput) Then
            SendKeys.Send("{End}")
            MessageBox.Show("Please input date in MM/DD/YYYY format.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub txtAbsentTo_Validated(sender As Object, e As EventArgs) Handles txtAbsentTo.Validated
        GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
    End Sub

    Private Sub txtEmployeeName_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtEmployeeName.Validating
        If String.IsNullOrEmpty(txtEmployeeName.Text.Trim) Then
            MessageBox.Show("Employee name is required.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub txtEmployeeScanId_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEmployeeScanId.KeyDown
        If e.KeyCode.Equals(Keys.Enter) Then
            e.Handled = True
            If String.IsNullOrEmpty(txtEmployeeScanId.Text.Trim) Then
                Me.ActiveControl = txtEmployeeScanId
                MessageBox.Show("Please enter employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            arrSplitted = Split(txtEmployeeScanId.Text.Trim, " ", 2)
            GetEmployeeInformation(arrSplitted(0).ToString)
        End If
    End Sub

#Region "Subroutines"

    Private Sub GetEmployeeInformation(employeeCode As String)
        Try
            Dim count As Integer = 0
            Dim prmCount(0) As SqlParameter
            prmCount(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
            prmCount(0).Value = employeeCode

            count = dbJeonsoft.ExecuteScalar("SELECT Count(Id) FROM viwGroupEmployees WHERE EmployeeCode = @EmployeeCode AND Active = 1", CommandType.Text, prmCount)

            If count > 0 Then 'direct
                Dim prmReader(0) As SqlParameter
                prmReader(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                prmReader(0).Value = employeeCode

                Using reader As IDataReader = dbLeaveFiling.ExecuteReader("RdEmployee", CommandType.StoredProcedure, prmReader)
                    While reader.Read
                        employeeId = reader.Item("Id")
                        departmentId = reader.Item("DepartmentId")
                        departmentName = reader.Item("DepartmentName")
                        positionId = reader.Item("PositionId")
                        positionName = reader.Item("PositionName")
                        txtEmployeeCode.Text = reader.Item("EmployeeCode").ToString.Trim
                        txtEmployeeName.Text = reader("EmployeeName").ToString.Trim

                        If reader.Item("TeamId") Is DBNull.Value Then
                            teamId = 0
                            teamName = String.Empty
                        Else
                            teamId = reader.Item("TeamId")
                            teamName = reader.Item("TeamName").ToString.Trim
                        End If
                    End While
                End Using

                cmbLeaveType.SelectedValue = 1
                cmbLeaveType.Enabled = True

                txtEmployeeScanId.Clear()
                txtEmployeeScanId.Enabled = False
                txtEmployeeName.Enabled = True
                txtEmployeeName.ReadOnly = True
                txtDate.Text = Format(dbLeaveFiling.GetServerDate, "MMMM dd, yyyy HH:mm")
                txtAbsentFrom.Enabled = True
                txtAbsentFrom.ReadOnly = False
                txtAbsentTo.Enabled = True
                txtAbsentTo.ReadOnly = False
                txtReason.Enabled = True
                txtReason.ReadOnly = False
                txtDiagnosis.Enabled = True
                txtDiagnosis.ReadOnly = False
                chkNotFtw.Enabled = True
                txtReason.Focus()

            Else 'agency (fmb)
                If employeeCode.Substring(0, 3).ToUpper.Trim.Equals("FMB") Then
                    employeeId = 0
                    txtEmployeeScanId.Clear()
                    txtEmployeeScanId.Enabled = False
                    txtEmployeeCode.Text = employeeCode
                    txtEmployeeCode.Text = StrConv(txtEmployeeCode.Text.Trim, VbStrConv.Uppercase)
                    txtEmployeeName.Enabled = True
                    txtEmployeeName.ReadOnly = False
                    txtDate.Text = Format(dbLeaveFiling.GetServerDate, "MMMM dd, yyyy HH:mm")
                    txtAbsentFrom.Enabled = True
                    txtAbsentFrom.ReadOnly = False
                    txtAbsentTo.Enabled = True
                    txtAbsentTo.ReadOnly = False
                    txtReason.Enabled = True
                    txtReason.ReadOnly = False
                    txtDiagnosis.Enabled = True
                    txtDiagnosis.ReadOnly = False
                    chkNotFtw.Enabled = True

                    cmbLeaveType.SelectedValue = 1
                    cmbLeaveType.Enabled = True

                    txtEmployeeName.Focus()
                Else
                    MessageBox.Show("Employee not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtEmployeeScanId.Focus()
                    Return
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'tag employee as `unfit to work` using shortcut key (F11)
    Private Sub NotFitToWork()
        Try
            If String.IsNullOrEmpty(txtEmployeeScanId.Text.Trim) AndAlso String.IsNullOrEmpty(txtEmployeeCode.Text.Trim) Then
                Me.ActiveControl = txtEmployeeScanId
                MessageBox.Show("Please enter employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If cmbLeaveType.SelectedValue = 0 Then
                MessageBox.Show("Please select leave type.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            If String.IsNullOrEmpty(txtReason.Text.Trim) Then
                MessageBox.Show("Please indicate the reason.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtReason
                Return
            End If

            If String.IsNullOrEmpty(txtDiagnosis.Text.Trim) Then
                MessageBox.Show("Please indicate the diagnosis.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtDiagnosis
                Return
            End If

            If CDate(txtAbsentFrom.Text).Date > CDate(txtAbsentTo.Text).Date Then
                MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = txtAbsentFrom
                Return
            End If

            'half day leaves
            If (cmbLeaveType.SelectedValue = 12 Or cmbLeaveType.SelectedIndex = 15 Or cmbLeaveType.SelectedValue = 16) AndAlso
                Not (CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date)) Then
                MessageBox.Show("Half-day leave should have the same dates.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.ActiveControl = cmbLeaveType
                Return
            End If

            SaveRecord(True)
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ResetForm()
        screenId = 0
        employeeId = 0
        departmentId = 0
        teamId = 0
        positionId = 0

        Me.Text = "New Record"
        txtEmployeeScanId.Enabled = True
        txtEmployeeScanId.Clear()
        txtEmployeeCode.Text = ""
        txtEmployeeName.Clear()
        txtEmployeeName.Enabled = False
        txtDate.Text = ""
        cmbLeaveType.Enabled = False
        cmbLeaveType.SelectedValue = 0
        txtAbsentFrom.Enabled = False
        txtAbsentFrom.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
        txtAbsentFrom.ValidatingType = GetType(System.DateTime)
        txtAbsentTo.Enabled = False
        txtAbsentTo.Text = String.Format("{0:MM/dd/yyyy}", GetLastWorkingDay(dbLeaveFiling.GetServerDate))
        txtAbsentTo.ValidatingType = GetType(System.DateTime)
        txtTotalDays.Text = 1
        txtReason.Clear()
        txtReason.Enabled = False
        txtDiagnosis.Clear()
        txtDiagnosis.Enabled = False
        chkNotFtw.Enabled = False
        chkNotFtw.CheckState = CheckState.Unchecked
        txtEmployeeScanId.Focus()
    End Sub

    Private Sub SaveRecord(isUnfitToWork As Boolean)
        Try
            Dim frmScreenList As frmScreenList = TryCast(Me.Owner, frmScreenList)

            If screenId = 0 Then 'new record
                Dim newScreeningRow As ScreeningRow = Me.dsLeaveFiling.Screening.NewScreeningRow

                If employeeId = 0 Then 'agency
                    Dim prmCntScreenDateRange(4) As SqlParameter 'check if has duplicate record in screening (date range)
                    prmCntScreenDateRange(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                    prmCntScreenDateRange(0).Value = Nothing
                    prmCntScreenDateRange(1) = New SqlParameter("@EmployeeCode", SqlDbType.NVarChar)
                    prmCntScreenDateRange(1).Value = txtEmployeeCode.Text.Trim
                    prmCntScreenDateRange(2) = New SqlParameter("@AbsentFrom", SqlDbType.Date)
                    prmCntScreenDateRange(2).Value = CDate(txtAbsentFrom.Text)
                    prmCntScreenDateRange(3) = New SqlParameter("@AbsentTo", SqlDbType.Date)
                    prmCntScreenDateRange(3).Value = CDate(txtAbsentTo.Text)
                    prmCntScreenDateRange(4) = New SqlParameter("@TotalCount", SqlDbType.Int)
                    prmCntScreenDateRange(4).Direction = ParameterDirection.Output

                    dbLeaveFiling.ExecuteScalar("CntScreeningDateRangeAgency", CommandType.StoredProcedure, prmCntScreenDateRange)

                    If prmCntScreenDateRange(4).Value > 0 Then 'do not allow duplicate entry in screening (date range) i.e. overlapping or in-between
                        MessageBox.Show("Record with the same date(s) already exists.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    Else
                        With newScreeningRow
                            .ScreenDate = dbLeaveFiling.GetServerDate
                            .ScreenBy = screenBy
                            .EmployeeId = 0
                            .EmployeeCode = txtEmployeeCode.Text.Trim
                            .EmployeeName = txtEmployeeName.Text.Trim
                            .AbsentFrom = CDate(txtAbsentFrom.Text)
                            .AbsentTo = CDate(txtAbsentTo.Text)
                            .LeaveTypeId = cmbLeaveType.SelectedValue
                            .Reason = txtReason.Text.Trim
                            .Diagnosis = txtDiagnosis.Text.Trim
                            .IsUsed = True 'set to used since the agency employess do not have access to leave application system
                            .SetModifiedByNull()
                            .SetModifiedDateNull()

                            Select Case cmbLeaveType.SelectedValue
                                Case 12, 15, 16
                                    .Quantity = 0.5
                                Case Else
                                    .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                            End Select

                            If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                .IsFitToWork = False
                            Else
                                If isUnfitToWork = True Then
                                    .IsFitToWork = False
                                Else
                                    .IsFitToWork = True
                                End If
                            End If
                        End With
                        Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                        Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                    End If

                Else 'direct
                    Dim prmCntScreenDateRange(4) As SqlParameter 'check if has duplicate record in screening (date range)
                    prmCntScreenDateRange(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                    prmCntScreenDateRange(0).Value = Nothing
                    prmCntScreenDateRange(1) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                    prmCntScreenDateRange(1).Value = employeeId
                    prmCntScreenDateRange(2) = New SqlParameter("@AbsentFrom", SqlDbType.Date)
                    prmCntScreenDateRange(2).Value = CDate(txtAbsentFrom.Text)
                    prmCntScreenDateRange(3) = New SqlParameter("@AbsentTo", SqlDbType.Date)
                    prmCntScreenDateRange(3).Value = CDate(txtAbsentTo.Text)
                    prmCntScreenDateRange(4) = New SqlParameter("@TotalCount", SqlDbType.Int)
                    prmCntScreenDateRange(4).Direction = ParameterDirection.Output

                    dbLeaveFiling.ExecuteScalar("CntScreeningDateRange", CommandType.StoredProcedure, prmCntScreenDateRange)

                    If prmCntScreenDateRange(4).Value > 0 Then 'do not allow duplicate entry in screening (date range) i.e. overlapping or in-between
                        MessageBox.Show("Record with the same date(s) already exists.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    Else
                        Dim prmCntLeaveDateExact(6) As SqlParameter 'check if has duplicate record in leave filing (exact date) i.e. leave filed in advance
                        prmCntLeaveDateExact(0) = New SqlParameter("@LeaveFileId", SqlDbType.Int)
                        prmCntLeaveDateExact(0).Value = Nothing
                        prmCntLeaveDateExact(1) = New SqlParameter("@ScreenId", SqlDbType.Int)
                        prmCntLeaveDateExact(1).Value = Nothing
                        prmCntLeaveDateExact(2) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                        prmCntLeaveDateExact(2).Value = employeeId
                        prmCntLeaveDateExact(3) = New SqlParameter("@StartDate", SqlDbType.Date)
                        prmCntLeaveDateExact(3).Value = CDate(txtAbsentFrom.Text)
                        prmCntLeaveDateExact(4) = New SqlParameter("@EndDate", SqlDbType.Date)
                        prmCntLeaveDateExact(4).Value = CDate(txtAbsentTo.Text)
                        prmCntLeaveDateExact(5) = New SqlParameter("@TotalLeaveFileId", SqlDbType.Int)
                        prmCntLeaveDateExact(5).Direction = ParameterDirection.Output
                        prmCntLeaveDateExact(6) = New SqlParameter("@TotalScreenId", SqlDbType.Int)
                        prmCntLeaveDateExact(6).Direction = ParameterDirection.Output

                        dbLeaveFiling.ExecuteScalar("CntLeaveFilingDateExact", CommandType.StoredProcedure, prmCntLeaveDateExact)

                        If prmCntLeaveDateExact(5).Value > 0 Then 'has duplicate record in leave filing (exact date), overwrite existing record
                            Dim rdrDateExact(2) As SqlParameter
                            rdrDateExact(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                            rdrDateExact(0).Value = employeeId
                            rdrDateExact(1) = New SqlParameter("@StartDate", SqlDbType.Date)
                            rdrDateExact(1).Value = CDate(txtAbsentFrom.Text)
                            rdrDateExact(2) = New SqlParameter("@EndDate", SqlDbType.Date)
                            rdrDateExact(2).Value = CDate(txtAbsentTo.Text)

                            Dim leaveFileId As Integer = 0
                            Dim startDate As Date = Nothing
                            Dim endDate As Date = Nothing
                            Dim question As String = String.Empty

                            Using rdrDate As IDataReader = dbLeaveFiling.ExecuteReader("RdLeaveFilingByLeaveDate", CommandType.StoredProcedure, rdrDateExact)
                                While rdrDate.Read
                                    leaveFileId = rdrDate.Item("LeaveFileId")
                                    startDate = CDate(rdrDate.Item("StartDate"))
                                    endDate = CDate(rdrDate.Item("EndDate"))
                                End While
                                rdrDate.Close()
                            End Using

                            If startDate.Equals(endDate) Then
                                question = String.Format("Employee has an existing leave dated {0}. Overwrite this record?", startDate.ToString("MMMM dd, yyyy"))
                            Else
                                question = String.Format("Employee has an existing leave dated from {0} to {1}. Overwrite this record?", startDate.ToString("MMMM dd, yyyy"),
                                                         endDate.ToString("MMMM dd, yyyy"))
                            End If

                            If MessageBox.Show(question, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                'direct - save to screening then proceed to automatic filing in leave application
                                'agency - save to screening only
                                If cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14 Then 'ecq leaves
                                    With newScreeningRow
                                        .ScreenDate = dbLeaveFiling.GetServerDate
                                        .ScreenBy = screenBy
                                        .EmployeeId = employeeId
                                        .EmployeeCode = txtEmployeeCode.Text.Trim
                                        .EmployeeName = txtEmployeeName.Text.Trim
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim
                                        .IsUsed = True
                                        .SetModifiedByNull()
                                        .SetModifiedDateNull()

                                        If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                            .IsFitToWork = False
                                        Else
                                            If isUnfitToWork = True Then
                                                .IsFitToWork = False
                                            Else
                                                .IsFitToWork = True
                                            End If
                                        End If
                                    End With
                                    Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                                    Me.adpLeaveFiling.FillByLeaveFileId(Me.dsLeaveFiling.LeaveFiling, leaveFileId)
                                    Dim leaveFilingRow As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.FindByLeaveFileId(leaveFileId)

                                    'overwrite the existing record in leave filing
                                    'change the modified date only, not the created date column
                                    'revert the status to pending if already processed by hr
                                    With leaveFilingRow
                                        .ScreenId = newScreeningRow.ScreenId
                                        .StartDate = CDate(txtAbsentFrom.Text)
                                        .EndDate = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                                        .Reason = txtReason.Text.Trim
                                        .LeaveCredits = GetLeaveCredits(employeeId)
                                        .LeaveBalance = GetLeaveBalance(employeeId)
                                        .ClinicIsApproved = True
                                        .ClinicId = screenBy
                                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                        .ClinicRemarks = txtDiagnosis.Text.Trim
                                        .IsLateFiling = 1
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .IsEncoded = False
                                        .IsDone = False

                                        .SuperiorIsApproved1 = 0
                                        .SetSuperiorApprovalDate1Null()
                                        .SetSuperiorRemarks1Null()

                                        .SuperiorIsApproved2 = 0
                                        .SetSuperiorApprovalDate2Null()
                                        .SetSuperiorRemarks2Null()

                                        .ManagerIsApproved = 0
                                        .SetManagerApprovalDateNull()
                                        .SetManagerRemarksNull()

                                        Dim prmCountRecipient(2) As SqlParameter
                                        prmCountRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                        prmCountRecipient(0).Value = departmentId
                                        prmCountRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                        prmCountRecipient(1).Value = teamId
                                        prmCountRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                        prmCountRecipient(2).Value = positionId

                                        'check if automatic filing has recipient
                                        Dim count As Integer = dbLeaveFiling.ExecuteScalar("SELECT COUNT(RecipientId) AS Count FROM dbo.Recipient " &
                                                                                           "WHERE DepartmentId = @DepartmentId AND TeamId = @TeamId " &
                                                                                           "AND PositionId = @PositionId ", CommandType.Text, prmCountRecipient)

                                        If count = 0 Then 'inform the developer regarding automatic filing with no recipients
                                            frmScreenList.SendDevNotif(employeeId, txtEmployeeName.Text.ToString.Trim, cmbLeaveType.SelectedValue,
                                                                       cmbLeaveType.Text, departmentId, departmentName, teamId, teamName, positionId, positionName)
                                            'set the last approver based on the majority of other records per department
                                            Dim managerId As Integer = 0
                                            Dim prmDeptId(0) As SqlParameter
                                            prmDeptId(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmDeptId(0).Value = departmentId

                                            managerId = dbLeaveFiling.ExecuteScalar("SELECT TOP 1 ManagerId FROM dbo.Recipient " &
                                                                                    "WHERE DepartmentId = @DepartmentId GROUP BY ManagerId " &
                                                                                    "ORDER BY COUNT(RecipientId) DESC",
                                                                                    CommandType.Text, prmDeptId)

                                            If employeeId = managerId Then 'employee is a manager, set dgm as the last approver
                                                .ManagerId = 70
                                            Else
                                                .ManagerId = managerId
                                            End If

                                            .RoutingStatusId = 3
                                        Else
                                            Dim prmRecipient(2) As SqlParameter
                                            prmRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmRecipient(0).Value = departmentId
                                            prmRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                            prmRecipient(1).Value = teamId
                                            prmRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                            prmRecipient(2).Value = positionId

                                            Using rdrRecipient As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, prmRecipient)
                                                While rdrRecipient.Read
                                                    If rdrRecipient.Item("SuperiorId1") Is DBNull.Value Then 'no superior 1
                                                        .SetSuperiorId1Null()

                                                        If rdrRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()

                                                            If employeeId = rdrRecipient.Item("ManagerId") Then 'employee is a manager, set dgm as the approver
                                                                .RoutingStatusId = 3
                                                                .ManagerId = 70 'dgm
                                                            Else
                                                                .RoutingStatusId = 3
                                                                .ManagerId = rdrRecipient.Item("ManagerId")
                                                            End If
                                                        Else
                                                            If employeeId = rdrRecipient.Item("SuperiorId2") Then
                                                                .RoutingStatusId = 3
                                                                .SetSuperiorId2Null()
                                                            Else
                                                                .RoutingStatusId = 4
                                                                .SuperiorId2 = rdrRecipient.Item("SuperiorId2")
                                                            End If
                                                        End If
                                                    Else 'with superior 1
                                                        If employeeId = rdrRecipient.Item("SuperiorId1") Then
                                                            .RoutingStatusId = 4
                                                            .SetSuperiorId1Null()
                                                        Else
                                                            .RoutingStatusId = 5
                                                            .SuperiorId1 = rdrRecipient.Item("SuperiorId1")
                                                        End If

                                                        If rdrRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()
                                                        Else
                                                            .SuperiorId2 = rdrRecipient.Item("SuperiorId2")
                                                        End If
                                                    End If

                                                    .ManagerId = rdrRecipient.Item("ManagerId")
                                                End While
                                                rdrRecipient.Close()
                                            End Using

                                            If .RoutingStatusId = 3 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                        .ManagerId,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                        .ManagerId,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                        CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                End If
                                            ElseIf .RoutingStatusId = 4 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                        .SuperiorId2,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                        .SuperiorId2,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                        CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                End If
                                            ElseIf .RoutingStatusId = 5 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                        .SuperiorId1,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                        .SuperiorId1,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                        CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                End If
                                            End If
                                        End If
                                    End With
                                Else 'other leave types
                                    With newScreeningRow
                                        .ScreenDate = dbLeaveFiling.GetServerDate
                                        .ScreenBy = screenBy
                                        .SetModifiedByNull()
                                        .SetModifiedDateNull()
                                        .EmployeeId = employeeId
                                        .EmployeeCode = txtEmployeeCode.Text.Trim
                                        .EmployeeName = txtEmployeeName.Text.Trim
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)

                                        Select Case cmbLeaveType.SelectedValue
                                            Case 12, 15, 16
                                                .Quantity = 0.5
                                            Case Else
                                                .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                                        End Select

                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim

                                        If isUnfitToWork = True Then
                                            .IsFitToWork = False
                                        Else
                                            .IsFitToWork = True
                                        End If

                                        .IsUsed = False
                                    End With
                                    Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                                    Me.adpLeaveFiling.FillByLeaveFileId(Me.dsLeaveFiling.LeaveFiling, leaveFileId)
                                    Dim leaveFilingRow As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.FindByLeaveFileId(leaveFileId)

                                    With leaveFilingRow
                                        .ScreenId = newScreeningRow.ScreenId
                                        .EmployeeId = employeeId
                                        .StartDate = CDate(txtAbsentFrom.Text)
                                        .EndDate = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .Reason = txtReason.Text.Trim
                                        .LeaveCredits = GetLeaveCredits(employeeId)
                                        .LeaveBalance = GetLeaveBalance(employeeId)
                                        .ClinicIsApproved = True
                                        .ClinicId = screenBy
                                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                        .ClinicRemarks = txtDiagnosis.Text.Trim
                                        .IsLateFiling = 1
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .IsEncoded = False
                                        .IsDone = False

                                        If .RoutingStatusId = 3 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .ManagerId,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .ManagerId,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            End If
                                        ElseIf .RoutingStatusId = 4 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId2,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId2,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            End If
                                        ElseIf .RoutingStatusId = 5 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId1,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId1,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            End If
                                        End If
                                    End With

                                    If lstLeaveTypeId.Contains(cmbLeaveType.SelectedValue) Then
                                        If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                            frmScreenList.SendRequestorNotif(employeeId,
                                                                                 CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                 cmbLeaveType.Text,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                 txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                 IIf(isUnfitToWork = True, "NO", "YES"))
                                        Else
                                            frmScreenList.SendRequestorNotif(employeeId,
                                                                                 CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                 cmbLeaveType.Text,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                 CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                 txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                 IIf(isUnfitToWork = True, "NO", "YES"))
                                        End If
                                    End If
                                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)
                                    Me.dsLeaveFiling.AcceptChanges()
                                End If
                            Else
                                Exit Sub
                            End If
                        Else
                            'no duplicate record in leave filing (exact date)
                            'check if has duplicate record in leave filing (date range) i.e. date selected is overlapped or in-between of an existing leave
                            Dim prmCountDate(6) As SqlParameter
                            prmCountDate(0) = New SqlParameter("@LeaveFileId", SqlDbType.Int)
                            prmCountDate(0).Value = Nothing
                            prmCountDate(1) = New SqlParameter("@ScreenId", SqlDbType.Int)
                            prmCountDate(1).Value = Nothing
                            prmCountDate(2) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                            prmCountDate(2).Value = employeeId
                            prmCountDate(3) = New SqlParameter("@StartDate", SqlDbType.Date)
                            prmCountDate(3).Value = CDate(txtAbsentFrom.Text)
                            prmCountDate(4) = New SqlParameter("@EndDate", SqlDbType.Date)
                            prmCountDate(4).Value = CDate(txtAbsentTo.Text)
                            prmCountDate(5) = New SqlParameter("TotalLeaveFileId", SqlDbType.Int)
                            prmCountDate(5).Direction = ParameterDirection.Output
                            prmCountDate(6) = New SqlParameter("TotalScreenId", SqlDbType.Int)
                            prmCountDate(6).Direction = ParameterDirection.Output

                            dbLeaveFiling.ExecuteScalar("CntLeaveFilingDateRange", CommandType.StoredProcedure, prmCountDate)

                            If prmCountDate(5).Value > 0 Then 'has duplicate entry in leave filing (date range)
                                Dim rdrDateExact(2) As SqlParameter
                                rdrDateExact(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                                rdrDateExact(0).Value = employeeId
                                rdrDateExact(1) = New SqlParameter("@StartDate", SqlDbType.Date)
                                rdrDateExact(1).Value = CDate(txtAbsentFrom.Text)
                                rdrDateExact(2) = New SqlParameter("@EndDate", SqlDbType.Date)
                                rdrDateExact(2).Value = CDate(txtAbsentTo.Text)

                                Dim leaveFileId As Integer = 0
                                Dim startDate As Date = Nothing
                                Dim endDate As Date = Nothing
                                Dim question2 As String = String.Empty

                                'get the dates of existing leave (date range)
                                Using rdrDate As IDataReader = dbLeaveFiling.ExecuteReader("RdLeaveFilingByLeaveDate", CommandType.StoredProcedure, rdrDateExact)
                                    While rdrDate.Read
                                        leaveFileId = rdrDate.Item("LeaveFileId")
                                        startDate = CDate(rdrDate.Item("StartDate"))
                                        endDate = CDate(rdrDate.Item("EndDate"))
                                    End While
                                    rdrDate.Close()
                                End Using

                                If startDate.Equals(endDate) Then
                                    question2 = String.Format("Employee has an existing leave dated {0}. Overwrite this record?", startDate.Date.ToString("MMMM dd, yyyy"))
                                Else
                                    question2 = String.Format("Employee has an existing leave dated from {0} to {1}. Overwrite this record?", startDate.Date.ToString("MMMM dd, yyyy"), endDate.Date.ToString("MMMM dd, yyyy"))
                                End If

                                If MessageBox.Show(question2, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                    Dim prmDel(0) As SqlParameter
                                    prmDel(0) = New SqlParameter("@LeaveFileId", SqlDbType.Int)
                                    prmDel(0).Value = leaveFileId
                                    dbLeaveFiling.ExecuteNonQuery("DELETE FROM dbo.LeaveFiling WHERE LeaveFileId = @LeaveFileId", CommandType.Text, prmDel)

                                    If cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14 Then
                                        'direct - save to screening then proceed to automatic filing in leave application
                                        'agency - save to screening only
                                        With newScreeningRow
                                            .ScreenDate = dbLeaveFiling.GetServerDate
                                            .ScreenBy = screenBy
                                            .EmployeeId = employeeId
                                            .EmployeeCode = txtEmployeeCode.Text.Trim
                                            .EmployeeName = txtEmployeeName.Text.Trim
                                            .AbsentFrom = CDate(txtAbsentFrom.Text)
                                            .AbsentTo = CDate(txtAbsentTo.Text)
                                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                            .LeaveTypeId = cmbLeaveType.SelectedValue
                                            .Reason = txtReason.Text.Trim
                                            .Diagnosis = txtDiagnosis.Text.Trim
                                            .IsUsed = True
                                            .SetModifiedByNull()
                                            .SetModifiedDateNull()

                                            If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                                .IsFitToWork = False
                                            Else
                                                If isUnfitToWork = True Then
                                                    .IsFitToWork = False
                                                Else
                                                    .IsFitToWork = True
                                                End If
                                            End If
                                        End With
                                        Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                                        Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                                        Dim newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                                        With newRowLeaveFiling
                                            .DateCreated = dbLeaveFiling.GetServerDate
                                            .ScreenId = newScreeningRow.ScreenId
                                            .EmployeeId = employeeId
                                            .DepartmentId = departmentId
                                            .TeamId = teamId
                                            .StartDate = CDate(txtAbsentFrom.Text)
                                            .EndDate = CDate(txtAbsentTo.Text)
                                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                            .Reason = txtReason.Text.Trim
                                            .LeaveCredits = GetLeaveCredits(employeeId)
                                            .LeaveBalance = GetLeaveBalance(employeeId)
                                            .ClinicIsApproved = 1
                                            .ClinicId = screenBy
                                            .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                            .ClinicRemarks = txtDiagnosis.Text.Trim
                                            .IsLateFiling = True
                                            .LeaveTypeId = cmbLeaveType.SelectedValue
                                            .SetModifiedByNull()
                                            .SetModifiedDateNull()
                                            .IsEncoded = False
                                            .IsDone = False

                                            .SuperiorIsApproved1 = 0
                                            .SetSuperiorApprovalDate1Null()
                                            .SetSuperiorRemarks1Null()

                                            .SuperiorIsApproved2 = 0
                                            .SetSuperiorApprovalDate2Null()
                                            .SetSuperiorRemarks2Null()

                                            .ManagerIsApproved = 0
                                            .SetManagerApprovalDateNull()
                                            .SetManagerRemarksNull()

                                            Dim prmCntRecipient(2) As SqlParameter
                                            prmCntRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmCntRecipient(0).Value = departmentId
                                            prmCntRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                            prmCntRecipient(1).Value = teamId
                                            prmCntRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                            prmCntRecipient(2).Value = positionId

                                            Dim count As Integer = dbLeaveFiling.ExecuteScalar("SELECT COUNT(RecipientId) AS Count FROM dbo.Recipient " &
                                                                                                   "WHERE DepartmentId = @DepartmentId AND TeamId = @TeamId AND " &
                                                                                                   "PositionId = @PositionId ", CommandType.Text, prmCntRecipient)

                                            If count = 0 Then 'inform the developer regarding automatic filing with no recipients
                                                frmScreenList.SendDevNotif(employeeId, txtEmployeeName.Text.ToString.Trim, cmbLeaveType.SelectedValue,
                                                                                     cmbLeaveType.Text, departmentId, departmentName, teamId, teamName, positionId, positionName)
                                                'set the last approver based on the majority of other records per department
                                                Dim managerId As Integer = 0
                                                Dim prmDeptId(0) As SqlParameter
                                                prmDeptId(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                                prmDeptId(0).Value = departmentId

                                                managerId = dbLeaveFiling.ExecuteScalar("SELECT TOP 1 ManagerId FROM dbo.Recipient " &
                                                                                            "WHERE DepartmentId = @DepartmentId GROUP BY ManagerId " &
                                                                                            "ORDER BY COUNT(RecipientId) DESC",
                                                                                            CommandType.Text, prmDeptId)

                                                If employeeId = managerId Then 'employee is a manager, set dgm as the approver
                                                    .ManagerId = 70
                                                Else
                                                    .ManagerId = managerId
                                                End If

                                                .RoutingStatusId = 3
                                            Else
                                                Dim prmRecipient(2) As SqlParameter
                                                prmRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                                prmRecipient(0).Value = departmentId
                                                prmRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                                prmRecipient(1).Value = teamId
                                                prmRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                                prmRecipient(2).Value = positionId

                                                Using readerRecipient As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, prmRecipient)
                                                    Dim superiorId1 As Integer = 0
                                                    Dim superiorId2 As Integer = 0
                                                    Dim managerId As Integer = 0

                                                    While readerRecipient.Read
                                                        If readerRecipient.Item("SuperiorId1") Is DBNull.Value Then 'no superior 1
                                                            .SetSuperiorId1Null()

                                                            If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                                .SetSuperiorId2Null()

                                                                If employeeId = readerRecipient.Item("ManagerId") Then 'employee is a manager, set dgm as the approver
                                                                    .RoutingStatusId = 3
                                                                    .ManagerId = 70 'dgm
                                                                Else
                                                                    .RoutingStatusId = 3
                                                                    .ManagerId = readerRecipient.Item("ManagerId")
                                                                End If
                                                            Else
                                                                If employeeId = readerRecipient.Item("SuperiorId2") Then
                                                                    .RoutingStatusId = 3
                                                                    .SetSuperiorId2Null()
                                                                Else
                                                                    .RoutingStatusId = 4
                                                                    .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                                End If
                                                            End If
                                                        Else 'with superior 1
                                                            If employeeId = readerRecipient.Item("SuperiorId1") Then
                                                                .RoutingStatusId = 4
                                                                .SetSuperiorId1Null()
                                                            Else
                                                                .RoutingStatusId = 5
                                                                .SuperiorId1 = readerRecipient.Item("SuperiorId1")
                                                            End If

                                                            If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                                .SetSuperiorId2Null()
                                                            Else
                                                                .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                            End If
                                                        End If

                                                        .ManagerId = readerRecipient.Item("ManagerId")
                                                        managerId = readerRecipient.Item("ManagerId")
                                                    End While
                                                    readerRecipient.Close()
                                                End Using
                                            End If
                                        End With
                                        Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(newRowLeaveFiling)
                                        Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                                        If newRowLeaveFiling.RoutingStatusId = 3 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                    newRowLeaveFiling.ManagerId,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName, CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                     newRowLeaveFiling.ManagerId,
                                                                                     cmbLeaveType.Text,
                                                                                     StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                     departmentName,
                                                                                     CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                     CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                     txtReason.Text.Trim)
                                            End If
                                        ElseIf newRowLeaveFiling.RoutingStatusId = 4 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                    newRowLeaveFiling.SuperiorId2,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                    newRowLeaveFiling.SuperiorId2,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            End If
                                        ElseIf newRowLeaveFiling.RoutingStatusId = 5 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                    newRowLeaveFiling.SuperiorId1,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                    newRowLeaveFiling.SuperiorId1,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                            End If
                                        End If
                                    Else 'other leave types
                                        With newScreeningRow
                                            .ScreenDate = dbLeaveFiling.GetServerDate
                                            .ScreenBy = screenBy
                                            .EmployeeId = employeeId
                                            .EmployeeCode = txtEmployeeCode.Text.Trim
                                            .EmployeeName = txtEmployeeName.Text.Trim
                                            .AbsentFrom = CDate(txtAbsentFrom.Text)
                                            .AbsentTo = CDate(txtAbsentTo.Text)
                                            .LeaveTypeId = cmbLeaveType.SelectedValue
                                            .Reason = txtReason.Text.Trim
                                            .Diagnosis = txtDiagnosis.Text.Trim
                                            .IsUsed = False
                                            .SetModifiedByNull()
                                            .SetModifiedDateNull()

                                            Select Case cmbLeaveType.SelectedValue
                                                Case 12, 15, 16
                                                    .Quantity = 0.5
                                                Case Else
                                                    .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                                            End Select

                                            If isUnfitToWork = True Then
                                                .IsFitToWork = False
                                            Else
                                                .IsFitToWork = True
                                            End If
                                        End With
                                        Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                                        Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                                        If lstLeaveTypeId.Contains(cmbLeaveType.SelectedValue) Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendRequestorNotif(employeeId,
                                                                                     CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                     cmbLeaveType.Text,
                                                                                     CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                     GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                     txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                     IIf(isUnfitToWork = True, "NO", "YES"))
                                            Else
                                                frmScreenList.SendRequestorNotif(employeeId,
                                                                                     CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                     cmbLeaveType.Text,
                                                                                     CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                     CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                     GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                     txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                     IIf(isUnfitToWork = True, "NO", "YES"))
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit Sub
                                End If
                            Else 'no existing record in leave filing, save record
                                If cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14 Then 'ecq leaves - automatic filing
                                    'direct - save to screening, automatic filing in leave application
                                    'agency - save to screening only
                                    With newScreeningRow
                                        .ScreenDate = dbLeaveFiling.GetServerDate
                                        .ScreenBy = screenBy
                                        .EmployeeId = employeeId
                                        .EmployeeCode = txtEmployeeCode.Text.Trim
                                        .EmployeeName = txtEmployeeName.Text.Trim
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim
                                        .IsUsed = True
                                        .SetModifiedByNull()
                                        .SetModifiedDateNull()

                                        If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                            .IsFitToWork = False
                                        Else
                                            If isUnfitToWork = True Then
                                                .IsFitToWork = False
                                            Else
                                                .IsFitToWork = True
                                            End If
                                        End If
                                    End With
                                    Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                                    Dim newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                                    With newRowLeaveFiling
                                        .DateCreated = dbLeaveFiling.GetServerDate
                                        .ScreenId = newScreeningRow.ScreenId
                                        .EmployeeId = employeeId
                                        .DepartmentId = departmentId
                                        .TeamId = teamId
                                        .StartDate = CDate(txtAbsentFrom.Text)
                                        .EndDate = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .Reason = txtReason.Text.Trim
                                        .LeaveCredits = GetLeaveCredits(employeeId)
                                        .LeaveBalance = GetLeaveBalance(employeeId)
                                        .ClinicIsApproved = 1
                                        .ClinicId = screenBy
                                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                        .ClinicRemarks = txtDiagnosis.Text.Trim
                                        .IsLateFiling = 1
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .SetModifiedByNull()
                                        .SetModifiedDateNull()
                                        .IsEncoded = 0
                                        .IsDone = 0

                                        .SuperiorIsApproved1 = 0
                                        .SetSuperiorApprovalDate1Null()
                                        .SetSuperiorRemarks1Null()

                                        .SuperiorIsApproved2 = 0
                                        .SetSuperiorApprovalDate2Null()
                                        .SetSuperiorRemarks2Null()

                                        .ManagerIsApproved = 0
                                        .SetManagerApprovalDateNull()
                                        .SetManagerRemarksNull()

                                        Dim prmCntRecipient(2) As SqlParameter
                                        prmCntRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                        prmCntRecipient(0).Value = departmentId
                                        prmCntRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                        prmCntRecipient(1).Value = teamId
                                        prmCntRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                        prmCntRecipient(2).Value = positionId

                                        Dim count As Integer = dbLeaveFiling.ExecuteScalar("SELECT COUNT(RecipientId) AS Count FROM dbo.Recipient " &
                                                                                               "WHERE DepartmentId = @DepartmentId AND TeamId = @TeamId AND " &
                                                                                               "PositionId = @PositionId ", CommandType.Text, prmCntRecipient)

                                        If count = 0 Then 'inform the developer regarding automatic filing with no recipients
                                            frmScreenList.SendDevNotif(employeeId, txtEmployeeName.Text.ToString.Trim, cmbLeaveType.SelectedValue,
                                                                           cmbLeaveType.Text, departmentId, departmentName, teamId, teamName, positionId, positionName)
                                            'set the last approver based on the majority of other records per department
                                            Dim managerId As Integer = 0
                                            Dim prmDeptId(0) As SqlParameter
                                            prmDeptId(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmDeptId(0).Value = departmentId

                                            managerId = dbLeaveFiling.ExecuteScalar("SELECT TOP 1 ManagerId FROM dbo.Recipient " &
                                                                                        "WHERE DepartmentId = @DepartmentId GROUP BY ManagerId " &
                                                                                        "ORDER BY COUNT(RecipientId) DESC",
                                                                                        CommandType.Text, prmDeptId)

                                            If employeeId = managerId Then 'employee is a manager, set dgm as the approver
                                                .ManagerId = 70
                                            Else
                                                .ManagerId = managerId
                                            End If

                                            .RoutingStatusId = 3
                                        Else
                                            Dim prmRecipient(2) As SqlParameter
                                            prmRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmRecipient(0).Value = departmentId
                                            prmRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                            prmRecipient(1).Value = teamId
                                            prmRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                            prmRecipient(2).Value = positionId

                                            Using readerRecipient As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, prmRecipient)
                                                Dim superiorId1 As Integer = 0
                                                Dim superiorId2 As Integer = 0
                                                Dim managerId As Integer = 0

                                                While readerRecipient.Read
                                                    If readerRecipient.Item("SuperiorId1") Is DBNull.Value Then 'no superior 1
                                                        .SetSuperiorId1Null()

                                                        If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()

                                                            If employeeId = readerRecipient.Item("ManagerId") Then 'employee is a manager, set dgm as the approver
                                                                .RoutingStatusId = 3
                                                                .ManagerId = 70 'dgm
                                                            Else
                                                                .RoutingStatusId = 3
                                                                .ManagerId = readerRecipient.Item("ManagerId")
                                                            End If
                                                        Else
                                                            If employeeId = readerRecipient.Item("SuperiorId2") Then
                                                                .RoutingStatusId = 3
                                                                .SetSuperiorId2Null()
                                                            Else
                                                                .RoutingStatusId = 4
                                                                .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                            End If
                                                        End If
                                                    Else 'with superior 1
                                                        If employeeId = readerRecipient.Item("SuperiorId1") Then
                                                            .RoutingStatusId = 4
                                                            .SetSuperiorId1Null()
                                                        Else
                                                            .RoutingStatusId = 5
                                                            .SuperiorId1 = readerRecipient.Item("SuperiorId1")
                                                        End If

                                                        If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()
                                                        Else
                                                            .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                        End If
                                                    End If

                                                    .ManagerId = readerRecipient.Item("ManagerId")
                                                    managerId = readerRecipient.Item("ManagerId")
                                                End While
                                                readerRecipient.Close()
                                            End Using
                                            Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(newRowLeaveFiling)
                                            Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                                            If newRowLeaveFiling.RoutingStatusId = 3 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                        newRowLeaveFiling.ManagerId,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName, CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                        newRowLeaveFiling.ManagerId,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                        CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                End If
                                            ElseIf newRowLeaveFiling.RoutingStatusId = 4 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                        newRowLeaveFiling.SuperiorId2,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                        newRowLeaveFiling.SuperiorId2,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                        CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                End If
                                            ElseIf newRowLeaveFiling.RoutingStatusId = 5 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                        newRowLeaveFiling.SuperiorId1,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                        newRowLeaveFiling.SuperiorId1,
                                                                                        cmbLeaveType.Text,
                                                                                        StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                        departmentName,
                                                                                        CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                        CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                        txtReason.Text.Trim)
                                                End If
                                            End If
                                        End If
                                    End With
                                Else 'other leave types
                                    With newScreeningRow
                                        .ScreenDate = dbLeaveFiling.GetServerDate
                                        .ScreenBy = screenBy
                                        .EmployeeId = employeeId
                                        .EmployeeCode = txtEmployeeCode.Text.Trim
                                        .EmployeeName = txtEmployeeName.Text.Trim
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim
                                        .IsUsed = False
                                        .SetModifiedByNull()
                                        .SetModifiedDateNull()

                                        Select Case cmbLeaveType.SelectedValue
                                            Case 12, 15, 16
                                                .Quantity = 0.5
                                            Case Else
                                                .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                                        End Select

                                        If isUnfitToWork = True Then
                                            .IsFitToWork = False
                                        Else
                                            .IsFitToWork = True
                                        End If
                                    End With
                                    Me.dsLeaveFiling.Screening.AddScreeningRow(newScreeningRow)
                                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)

                                    If lstLeaveTypeId.Contains(cmbLeaveType.SelectedValue) Then
                                        If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                            frmScreenList.SendRequestorNotif(employeeId,
                                                                                     CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                     cmbLeaveType.Text,
                                                                                     CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                     GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                     txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                     IIf(isUnfitToWork = True, "NO", "YES"))
                                        Else
                                            frmScreenList.SendRequestorNotif(employeeId,
                                                                                     CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                     cmbLeaveType.Text,
                                                                                     CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                     CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                     GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                     txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                     IIf(isUnfitToWork = True, "NO", "YES"))
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                Me.dsLeaveFiling.AcceptChanges()
                frmScreenList.RefreshList()
                ResetForm()

            Else 'old record
                Dim rowScreening As ScreeningRow = Me.dsLeaveFiling.Screening.FindByScreenId(screenId)

                If employeeId = 0 Then 'agency
                    Dim prmCntScreenDateRange(4) As SqlParameter 'check if has duplicate record in screening (date range)
                    prmCntScreenDateRange(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                    prmCntScreenDateRange(0).Value = screenId
                    prmCntScreenDateRange(1) = New SqlParameter("@EmployeeCode", SqlDbType.NVarChar)
                    prmCntScreenDateRange(1).Value = txtEmployeeCode.Text.Trim
                    prmCntScreenDateRange(2) = New SqlParameter("@AbsentFrom", SqlDbType.Date)
                    prmCntScreenDateRange(2).Value = CDate(txtAbsentFrom.Text)
                    prmCntScreenDateRange(3) = New SqlParameter("@AbsentTo", SqlDbType.Date)
                    prmCntScreenDateRange(3).Value = CDate(txtAbsentTo.Text)
                    prmCntScreenDateRange(4) = New SqlParameter("TotalCount", SqlDbType.Int)
                    prmCntScreenDateRange(4).Direction = ParameterDirection.Output

                    dbLeaveFiling.ExecuteScalar("CntScreeningDateRangeAgency", CommandType.StoredProcedure, prmCntScreenDateRange)

                    If prmCntScreenDateRange(4).Value > 0 Then 'do not allow duplicate entry in screening (date range) i.e. overlapping or in-between
                        MessageBox.Show("Record with the same date(s) already exists.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    Else
                        With rowScreening
                            .ModifiedBy = screenBy
                            .ModifiedDate = dbLeaveFiling.GetServerDate
                            .EmployeeId = employeeId
                            .EmployeeCode = txtEmployeeCode.Text.Trim
                            .EmployeeName = txtEmployeeName.Text.Trim
                            .AbsentFrom = CDate(txtAbsentFrom.Text)
                            .AbsentTo = CDate(txtAbsentTo.Text)
                            .LeaveTypeId = cmbLeaveType.SelectedValue
                            .Reason = txtReason.Text.Trim
                            .Diagnosis = txtDiagnosis.Text.Trim

                            If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                .IsFitToWork = False
                            Else
                                If isUnfitToWork = True Then
                                    .IsFitToWork = False
                                Else
                                    .IsFitToWork = True
                                End If
                            End If

                            Select Case cmbLeaveType.SelectedValue
                                Case 12, 15, 16
                                    .Quantity = 0.5
                                Case Else
                                    .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                            End Select
                        End With
                    End If
                Else 'direct
                    Dim prmCntScreenDateRange(4) As SqlParameter 'check if has duplicate record in screening (date range)
                    prmCntScreenDateRange(0) = New SqlParameter("@ScreenId", SqlDbType.Int)
                    prmCntScreenDateRange(0).Value = screenId
                    prmCntScreenDateRange(1) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                    prmCntScreenDateRange(1).Value = employeeId
                    prmCntScreenDateRange(2) = New SqlParameter("@AbsentFrom", SqlDbType.Date)
                    prmCntScreenDateRange(2).Value = CDate(txtAbsentFrom.Text)
                    prmCntScreenDateRange(3) = New SqlParameter("@AbsentTo", SqlDbType.Date)
                    prmCntScreenDateRange(3).Value = CDate(txtAbsentTo.Text)
                    prmCntScreenDateRange(4) = New SqlParameter("TotalCount", SqlDbType.Int)
                    prmCntScreenDateRange(4).Direction = ParameterDirection.Output

                    dbLeaveFiling.ExecuteScalar("CntScreeningDateRange", CommandType.StoredProcedure, prmCntScreenDateRange)

                    If prmCntScreenDateRange(4).Value > 0 Then 'do not allow duplicate entry in screening (date range) i.e. overlapping or in-between
                        MessageBox.Show("Record with the same date(s) already exists.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    Else
                        Dim prmCntLeaveDateExact(6) As SqlParameter  'check if has duplicate record in leave filing (exact date) but not the same screen id i.e. leave filed in advance
                        prmCntLeaveDateExact(0) = New SqlParameter("@LeaveFileId", SqlDbType.Int)
                        prmCntLeaveDateExact(0).Value = Nothing
                        prmCntLeaveDateExact(1) = New SqlParameter("@ScreenId", SqlDbType.Int)
                        prmCntLeaveDateExact(1).Value = screenId
                        prmCntLeaveDateExact(2) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                        prmCntLeaveDateExact(2).Value = employeeId
                        prmCntLeaveDateExact(3) = New SqlParameter("@StartDate", SqlDbType.Date)
                        prmCntLeaveDateExact(3).Value = CDate(txtAbsentFrom.Text)
                        prmCntLeaveDateExact(4) = New SqlParameter("@EndDate", SqlDbType.Date)
                        prmCntLeaveDateExact(4).Value = CDate(txtAbsentTo.Text)
                        prmCntLeaveDateExact(5) = New SqlParameter("TotalLeaveFileId", SqlDbType.Int)
                        prmCntLeaveDateExact(5).Direction = ParameterDirection.Output
                        prmCntLeaveDateExact(6) = New SqlParameter("TotalScreenId", SqlDbType.Int)
                        prmCntLeaveDateExact(6).Direction = ParameterDirection.Output

                        dbLeaveFiling.ExecuteScalar("CntLeaveFilingDateExact", CommandType.StoredProcedure, prmCntLeaveDateExact)

                        If prmCntLeaveDateExact(5).Value > 0 Then 'has duplicate record in leave filing (exact date), overwrite existing record
                            Dim leaveFileId As Integer = 0 'get existing record in leave filing (exact date)
                            Dim startDate As Date = Nothing
                            Dim endDate As Date = Nothing
                            Dim question As String = String.Empty

                            Dim rdrDateExact(2) As SqlParameter
                            rdrDateExact(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                            rdrDateExact(0).Value = employeeId
                            rdrDateExact(1) = New SqlParameter("@StartDate", SqlDbType.Date)
                            rdrDateExact(1).Value = CDate(txtAbsentFrom.Text)
                            rdrDateExact(2) = New SqlParameter("@EndDate", SqlDbType.Date)
                            rdrDateExact(2).Value = CDate(txtAbsentTo.Text)

                            Using rdrDate As IDataReader = dbLeaveFiling.ExecuteReader("RdLeaveFilingByLeaveDate", CommandType.StoredProcedure, rdrDateExact)
                                While rdrDate.Read
                                    leaveFileId = rdrDate.Item("LeaveFileId")
                                    startDate = CDate(rdrDate.Item("StartDate"))
                                    endDate = CDate(rdrDate.Item("EndDate"))
                                End While
                                rdrDate.Close()
                            End Using

                            If startDate.Equals(endDate) Then
                                question = String.Format("Employee has an existing leave dated {0}. Overwrite this record?", startDate.ToString("MMMM dd, yyyy"))
                            Else
                                question = String.Format("Employee has an existing leave dated from {0} to {1}. Overwrite this record?", startDate.ToString("MMMM dd, yyyy"),
                                                     endDate.ToString("MMMM dd, yyyy"))
                            End If

                            If MessageBox.Show(question, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                'direct - update the screening record and leave filing record
                                'agency - update screening record only
                                If cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14 Then 'ecq leaves
                                    With rowScreening
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .EmployeeId = employeeId
                                        .EmployeeCode = txtEmployeeCode.Text.Trim
                                        .EmployeeName = txtEmployeeName.Text.Trim
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim

                                        If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                            .IsFitToWork = False
                                        Else
                                            If isUnfitToWork = True Then
                                                .IsFitToWork = False
                                            Else
                                                .IsFitToWork = True
                                            End If
                                        End If
                                    End With

                                    Me.adpLeaveFiling.FillByLeaveFileId(Me.dsLeaveFiling.LeaveFiling, leaveFileId)
                                    Dim rowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.FindByLeaveFileId(leaveFileId)

                                    'overwrite existing record in leave filing
                                    'change the modified date only, not the date created column
                                    'revert status to pending if already processed by hr
                                    With rowLeaveFiling
                                        .ScreenId = screenId
                                        .StartDate = CDate(txtAbsentFrom.Text)
                                        .EndDate = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .Reason = txtReason.Text.Trim
                                        .LeaveCredits = GetLeaveCredits(employeeId)
                                        .LeaveBalance = GetLeaveBalance(employeeId)
                                        .ClinicIsApproved = True
                                        .ClinicId = screenBy
                                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                        .ClinicRemarks = txtDiagnosis.Text.Trim
                                        .IsLateFiling = 1
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .IsEncoded = False
                                        .IsDone = False

                                        .SuperiorIsApproved1 = 0
                                        .SetSuperiorApprovalDate1Null()
                                        .SetSuperiorRemarks1Null()

                                        .SuperiorIsApproved2 = 0
                                        .SetSuperiorApprovalDate2Null()
                                        .SetSuperiorRemarks2Null()

                                        .ManagerIsApproved = 0
                                        .SetManagerApprovalDateNull()
                                        .SetManagerRemarksNull()

                                        Dim prmCountRecipient(2) As SqlParameter
                                        prmCountRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                        prmCountRecipient(0).Value = departmentId
                                        prmCountRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                        prmCountRecipient(1).Value = teamId
                                        prmCountRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                        prmCountRecipient(2).Value = positionId

                                        'check if automatic filing has recipient
                                        Dim count As Integer = dbLeaveFiling.ExecuteScalar("SELECT COUNT(RecipientId) AS Count FROM dbo.Recipient " &
                                                                                           "WHERE DepartmentId = @DepartmentId AND TeamId = @TeamId And PositionId = @PositionId ",
                                                                                           CommandType.Text, prmCountRecipient)

                                        If count = 0 Then 'inform the developer regarding automatic filing with no recipients
                                            frmScreenList.SendDevNotif(employeeId, txtEmployeeName.Text.ToString.Trim, cmbLeaveType.SelectedValue,
                                                                       cmbLeaveType.Text, departmentId, departmentName, teamId, teamName, positionId, positionName)
                                            'set the last approver based on the majority of other records per department
                                            Dim managerId As Integer = 0
                                            Dim prmDeptId(0) As SqlParameter
                                            prmDeptId(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmDeptId(0).Value = departmentId

                                            managerId = dbLeaveFiling.ExecuteScalar("SELECT TOP 1 ManagerId FROM dbo.Recipient " &
                                                                                    "WHERE DepartmentId = @DepartmentId GROUP BY ManagerId " &
                                                                                    "ORDER BY COUNT(RecipientId) DESC",
                                                                                    CommandType.Text, prmDeptId)

                                            If employeeId = managerId Then 'employee is a manager, set dgm as the approver
                                                .ManagerId = 70
                                            Else
                                                .ManagerId = managerId
                                            End If

                                            .RoutingStatusId = 3
                                        Else
                                            Dim prmRecipient(2) As SqlParameter
                                            prmRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmRecipient(0).Value = departmentId
                                            prmRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                            prmRecipient(1).Value = teamId
                                            prmRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                            prmRecipient(2).Value = positionId

                                            Using rdrRecipient As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, prmRecipient)
                                                While rdrRecipient.Read
                                                    If rdrRecipient.Item("SuperiorId1") Is DBNull.Value Then 'no superior 1
                                                        .SetSuperiorId1Null()

                                                        If rdrRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()

                                                            If employeeId = rdrRecipient.Item("ManagerId") Then 'employee is a manager, set dgm as the approver
                                                                .RoutingStatusId = 3
                                                                .ManagerId = 70 'dgm
                                                            Else
                                                                .RoutingStatusId = 3
                                                                .ManagerId = rdrRecipient.Item("ManagerId")
                                                            End If
                                                        Else
                                                            If employeeId = rdrRecipient.Item("SuperiorId2") Then
                                                                .RoutingStatusId = 3
                                                                .SetSuperiorId2Null()
                                                            Else
                                                                .RoutingStatusId = 4
                                                                .SuperiorId2 = rdrRecipient.Item("SuperiorId2")
                                                            End If
                                                        End If
                                                    Else 'with superior 1
                                                        If employeeId = rdrRecipient.Item("SuperiorId1") Then
                                                            .RoutingStatusId = 4
                                                            .SetSuperiorId1Null()
                                                        Else
                                                            .RoutingStatusId = 5
                                                            .SuperiorId1 = rdrRecipient.Item("SuperiorId1")
                                                        End If

                                                        If rdrRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()
                                                        Else
                                                            .SuperiorId2 = rdrRecipient.Item("SuperiorId2")
                                                        End If
                                                    End If

                                                    .ManagerId = rdrRecipient.Item("ManagerId")
                                                End While
                                                rdrRecipient.Close()
                                            End Using

                                            If .RoutingStatusId = 3 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .ManagerId,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .ManagerId,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase), departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                                End If
                                            ElseIf .RoutingStatusId = 4 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId2,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId2,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                                End If
                                            ElseIf .RoutingStatusId = 5 Then
                                                If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId1,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                                Else
                                                    frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                    .SuperiorId1,
                                                                                    cmbLeaveType.Text,
                                                                                    StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                    departmentName,
                                                                                    CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                    CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                    txtReason.Text.Trim)
                                                End If
                                            End If
                                        End If
                                    End With
                                Else 'other leave types
                                    With rowScreening
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .EmployeeId = employeeId
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim

                                        Select Case cmbLeaveType.SelectedValue
                                            Case 12, 15, 16
                                                .Quantity = 0.5
                                            Case Else
                                                .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        End Select

                                        If isUnfitToWork = True Then
                                            .IsFitToWork = False
                                        Else
                                            .IsFitToWork = True
                                        End If
                                    End With

                                    Me.adpLeaveFiling.FillByLeaveFileId(Me.dsLeaveFiling.LeaveFiling, leaveFileId)
                                    Dim leaveFilingRow As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.FindByLeaveFileId(leaveFileId)

                                    With leaveFilingRow
                                        .ScreenId = screenId
                                        .StartDate = CDate(txtAbsentFrom.Text)
                                        .EndDate = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .Reason = txtReason.Text.Trim
                                        .LeaveCredits = GetLeaveCredits(employeeId)
                                        .LeaveBalance = GetLeaveBalance(employeeId)
                                        .ClinicIsApproved = True
                                        .ClinicId = screenBy
                                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                        .ClinicRemarks = txtDiagnosis.Text.Trim
                                        .IsLateFiling = 1
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .IsEncoded = False
                                        .IsDone = False

                                        If .RoutingStatusId = 3 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                .ManagerId,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                .ManagerId,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            End If
                                        ElseIf .RoutingStatusId = 4 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                .SuperiorId2,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                .SuperiorId2,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            End If
                                        ElseIf .RoutingStatusId = 5 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                .SuperiorId1,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(.LeaveFileId,
                                                                                .SuperiorId1,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            End If
                                        End If
                                    End With

                                    If lstLeaveTypeId.Contains(cmbLeaveType.SelectedValue) Then
                                        If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                            frmScreenList.SendRequestorNotif(employeeId,
                                                                                 CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                 cmbLeaveType.Text,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                 txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                 IIf(isUnfitToWork = True, "NO", "YES"))
                                        Else
                                            frmScreenList.SendRequestorNotif(employeeId,
                                                                                 CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                 cmbLeaveType.Text,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                 CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                 txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                 IIf(isUnfitToWork = True, "NO", "YES"))
                                        End If
                                    End If
                                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)
                                    Me.dsLeaveFiling.AcceptChanges()
                                End If
                            Else
                                Exit Sub
                            End If
                        Else
                            'no duplicate record in leave filing (exact date)
                            'check if has duplicate record in leave filing (date range) i.e. date selected is overlapped or in-between of existing leave
                            Dim prmCountDate(6) As SqlParameter
                            prmCountDate(0) = New SqlParameter("@LeaveFileId", SqlDbType.Int)
                            prmCountDate(0).Value = Nothing
                            prmCountDate(1) = New SqlParameter("@ScreenId", SqlDbType.Int)
                            prmCountDate(1).Value = screenId
                            prmCountDate(2) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                            prmCountDate(2).Value = employeeId
                            prmCountDate(3) = New SqlParameter("@StartDate", SqlDbType.Date)
                            prmCountDate(3).Value = CDate(txtAbsentFrom.Text)
                            prmCountDate(4) = New SqlParameter("@EndDate", SqlDbType.Date)
                            prmCountDate(4).Value = CDate(txtAbsentTo.Text)
                            prmCountDate(5) = New SqlParameter("TotalLeaveFileId", SqlDbType.Int)
                            prmCountDate(5).Direction = ParameterDirection.Output
                            prmCountDate(6) = New SqlParameter("TotalScreenId", SqlDbType.Int)
                            prmCountDate(6).Direction = ParameterDirection.Output

                            dbLeaveFiling.ExecuteScalar("CntLeaveFilingDateRange", CommandType.StoredProcedure, prmCountDate)

                            If prmCountDate(5).Value > 0 Then 'has duplicate entry in leave filing
                                Dim rdrDateExact(2) As SqlParameter
                                rdrDateExact(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                                rdrDateExact(0).Value = employeeId
                                rdrDateExact(1) = New SqlParameter("@StartDate", SqlDbType.Date)
                                rdrDateExact(1).Value = CDate(txtAbsentFrom.Text)
                                rdrDateExact(2) = New SqlParameter("@EndDate", SqlDbType.Date)
                                rdrDateExact(2).Value = CDate(txtAbsentTo.Text)

                                Dim leaveFileId As Integer = 0
                                Dim startDate As Date = Nothing
                                Dim endDate As Date = Nothing
                                Dim question As String = String.Empty

                                'get the dates of existing leave (date range)
                                Using rdrDate As IDataReader = dbLeaveFiling.ExecuteReader("RdLeaveFilingByLeaveDate", CommandType.StoredProcedure, rdrDateExact)
                                    While rdrDate.Read
                                        leaveFileId = rdrDate.Item("LeaveFileId")
                                        startDate = CDate(rdrDate.Item("StartDate"))
                                        endDate = CDate(rdrDate.Item("EndDate"))
                                    End While
                                    rdrDate.Close()
                                End Using

                                If startDate.Equals(endDate) Then
                                    question = String.Format("Employee has an existing leave dated {0}. Overwrite this record?", startDate.Date.ToString("MMMM dd, yyyy"))
                                Else
                                    question = String.Format("Employee has an existing leave dated from {0} to {1}. Overwrite this record?", startDate.Date.ToString("MMMM dd, yyyy"), endDate.Date.ToString("MMMM dd, yyyy"))
                                End If

                                If MessageBox.Show(question, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                    Dim prmDel(0) As SqlParameter
                                    prmDel(0) = New SqlParameter("@LeaveFileId", SqlDbType.Int)
                                    prmDel(0).Value = leaveFileId
                                    dbLeaveFiling.ExecuteNonQuery("DELETE FROM dbo.LeaveFiling WHERE LeaveFileId = @LeaveFileId", CommandType.Text, prmDel)

                                    If cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14 Then
                                        'direct - save to screening then proceed to automatic filing in leave application
                                        'agency - save to screening only
                                        With rowScreening
                                            .ScreenDate = dbLeaveFiling.GetServerDate
                                            .ScreenBy = screenBy
                                            .EmployeeId = employeeId
                                            .EmployeeCode = txtEmployeeCode.Text.Trim
                                            .EmployeeName = txtEmployeeName.Text.Trim
                                            .AbsentFrom = CDate(txtAbsentFrom.Text)
                                            .AbsentTo = CDate(txtAbsentTo.Text)
                                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                            .LeaveTypeId = cmbLeaveType.SelectedValue
                                            .Reason = txtReason.Text.Trim
                                            .Diagnosis = txtDiagnosis.Text.Trim
                                            .IsUsed = True
                                            .ModifiedBy = screenBy
                                            .ModifiedDate = dbLeaveFiling.GetServerDate

                                            If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                                .IsFitToWork = False
                                            Else
                                                If isUnfitToWork = True Then
                                                    .IsFitToWork = False
                                                Else
                                                    .IsFitToWork = True
                                                End If
                                            End If
                                        End With

                                        Dim newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                                        With newRowLeaveFiling
                                            .DateCreated = dbLeaveFiling.GetServerDate
                                            .ScreenId = rowScreening.ScreenId
                                            .EmployeeId = employeeId
                                            .DepartmentId = departmentId
                                            .TeamId = teamId
                                            .StartDate = CDate(txtAbsentFrom.Text)
                                            .EndDate = CDate(txtAbsentTo.Text)
                                            .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                            .Reason = txtReason.Text.Trim
                                            .LeaveCredits = GetLeaveCredits(employeeId)
                                            .LeaveBalance = GetLeaveBalance(employeeId)
                                            .ClinicIsApproved = 1
                                            .ClinicId = screenBy
                                            .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                            .ClinicRemarks = txtDiagnosis.Text.Trim
                                            .IsLateFiling = True
                                            .LeaveTypeId = cmbLeaveType.SelectedValue
                                            .SetModifiedByNull()
                                            .SetModifiedDateNull()
                                            .IsEncoded = False
                                            .IsDone = False

                                            .SuperiorIsApproved1 = 0
                                            .SetSuperiorApprovalDate1Null()
                                            .SetSuperiorRemarks1Null()

                                            .SuperiorIsApproved2 = 0
                                            .SetSuperiorApprovalDate2Null()
                                            .SetSuperiorRemarks2Null()

                                            .ManagerIsApproved = 0
                                            .SetManagerApprovalDateNull()
                                            .SetManagerRemarksNull()

                                            Dim prmCntRecipient(2) As SqlParameter
                                            prmCntRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmCntRecipient(0).Value = departmentId
                                            prmCntRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                            prmCntRecipient(1).Value = teamId
                                            prmCntRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                            prmCntRecipient(2).Value = positionId

                                            Dim count As Integer = dbLeaveFiling.ExecuteScalar("SELECT COUNT(RecipientId) AS Count FROM dbo.Recipient " &
                                                                                               "WHERE DepartmentId = @DepartmentId AND TeamId = @TeamId AND " &
                                                                                               "PositionId = @PositionId ", CommandType.Text, prmCntRecipient)

                                            If count = 0 Then 'inform the developer regarding automatic filing with no recipients
                                                frmScreenList.SendDevNotif(employeeId, txtEmployeeName.Text.ToString.Trim, cmbLeaveType.SelectedValue,
                                                                                 cmbLeaveType.Text, departmentId, departmentName, teamId, teamName, positionId, positionName)
                                                'set the last approver based on the majority of other records per department
                                                Dim managerId As Integer = 0
                                                Dim prmDeptId(0) As SqlParameter
                                                prmDeptId(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                                prmDeptId(0).Value = departmentId

                                                managerId = dbLeaveFiling.ExecuteScalar("SELECT TOP 1 ManagerId FROM dbo.Recipient " &
                                                                                        "WHERE DepartmentId = @DepartmentId GROUP BY ManagerId " &
                                                                                        "ORDER BY COUNT(RecipientId) DESC",
                                                                                        CommandType.Text, prmDeptId)

                                                If employeeId = managerId Then 'employee is a manager, set dgm as the approver
                                                    .ManagerId = 70
                                                Else
                                                    .ManagerId = managerId
                                                End If

                                                .RoutingStatusId = 3
                                            Else
                                                Dim prmRecipient(2) As SqlParameter
                                                prmRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                                prmRecipient(0).Value = departmentId
                                                prmRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                                prmRecipient(1).Value = teamId
                                                prmRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                                prmRecipient(2).Value = positionId

                                                Using readerRecipient As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, prmRecipient)
                                                    Dim superiorId1 As Integer = 0
                                                    Dim superiorId2 As Integer = 0
                                                    Dim managerId As Integer = 0

                                                    While readerRecipient.Read
                                                        If readerRecipient.Item("SuperiorId1") Is DBNull.Value Then 'no superior 1
                                                            .SetSuperiorId1Null()

                                                            If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                                .SetSuperiorId2Null()

                                                                If employeeId = readerRecipient.Item("ManagerId") Then 'employee is a manager, set dgm as the approver
                                                                    .RoutingStatusId = 3
                                                                    .ManagerId = 70 'dgm
                                                                Else
                                                                    .RoutingStatusId = 3
                                                                    .ManagerId = readerRecipient.Item("ManagerId")
                                                                End If
                                                            Else
                                                                If employeeId = readerRecipient.Item("SuperiorId2") Then
                                                                    .RoutingStatusId = 3
                                                                    .SetSuperiorId2Null()
                                                                Else
                                                                    .RoutingStatusId = 4
                                                                    .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                                End If
                                                            End If
                                                        Else 'with superior 1
                                                            If employeeId = readerRecipient.Item("SuperiorId1") Then
                                                                .RoutingStatusId = 4
                                                                .SetSuperiorId1Null()
                                                            Else
                                                                .RoutingStatusId = 5
                                                                .SuperiorId1 = readerRecipient.Item("SuperiorId1")
                                                            End If

                                                            If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                                .SetSuperiorId2Null()
                                                            Else
                                                                .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                            End If
                                                        End If

                                                        .ManagerId = readerRecipient.Item("ManagerId")
                                                        managerId = readerRecipient.Item("ManagerId")
                                                    End While
                                                    readerRecipient.Close()
                                                End Using
                                            End If
                                        End With
                                        Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(newRowLeaveFiling)
                                        Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                                        If newRowLeaveFiling.RoutingStatusId = 3 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                newRowLeaveFiling.ManagerId,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName, CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                 newRowLeaveFiling.ManagerId,
                                                                                 cmbLeaveType.Text,
                                                                                 StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                 departmentName,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                 CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 txtReason.Text.Trim)
                                            End If
                                        ElseIf newRowLeaveFiling.RoutingStatusId = 4 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                newRowLeaveFiling.SuperiorId2,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                newRowLeaveFiling.SuperiorId2,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            End If
                                        ElseIf newRowLeaveFiling.RoutingStatusId = 5 Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                newRowLeaveFiling.SuperiorId1,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            Else
                                                frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                                newRowLeaveFiling.SuperiorId1,
                                                                                cmbLeaveType.Text,
                                                                                StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                                departmentName,
                                                                                CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                txtReason.Text.Trim)
                                            End If
                                        End If
                                    Else 'other leave types
                                        With rowScreening
                                            .ScreenDate = dbLeaveFiling.GetServerDate
                                            .ScreenBy = screenBy
                                            .EmployeeId = employeeId
                                            .EmployeeCode = txtEmployeeCode.Text.Trim
                                            .EmployeeName = txtEmployeeName.Text.Trim
                                            .AbsentFrom = CDate(txtAbsentFrom.Text)
                                            .AbsentTo = CDate(txtAbsentTo.Text)
                                            .LeaveTypeId = cmbLeaveType.SelectedValue
                                            .Reason = txtReason.Text.Trim
                                            .Diagnosis = txtDiagnosis.Text.Trim
                                            .ModifiedBy = screenBy
                                            .ModifiedDate = dbLeaveFiling.GetServerDate

                                            Select Case cmbLeaveType.SelectedValue
                                                Case 12, 15, 16
                                                    .Quantity = 0.5
                                                Case Else
                                                    .Quantity = GetTotalDays(txtAbsentFrom.Text, txtAbsentTo.Text)
                                            End Select

                                            If isUnfitToWork = True Then
                                                .IsFitToWork = False
                                            Else
                                                .IsFitToWork = True
                                            End If
                                        End With

                                        If lstLeaveTypeId.Contains(cmbLeaveType.SelectedValue) Then
                                            If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                                frmScreenList.SendRequestorNotif(employeeId,
                                                                                 CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                 cmbLeaveType.Text,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                 txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                 IIf(isUnfitToWork = True, "NO", "YES"))
                                            Else
                                                frmScreenList.SendRequestorNotif(employeeId,
                                                                                 CDate(dbLeaveFiling.GetServerDate).ToString("MMMM dd, yyyy hh:mm tt"),
                                                                                 cmbLeaveType.Text,
                                                                                 CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                                 CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                                 GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text)),
                                                                                 txtReason.Text.Trim, txtDiagnosis.Text.Trim,
                                                                                 IIf(isUnfitToWork = True, "NO", "YES"))
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit Sub
                                End If
                            Else 'no existing leave (leave filing), save record
                                If cmbLeaveType.SelectedValue = 9 Or cmbLeaveType.SelectedValue = 14 Then
                                    'direct - save to screening, automatic filing in leave application
                                    'agency - save to screening only
                                    With rowScreening
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .EmployeeId = employeeId
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim
                                        .IsUsed = True

                                        If cmbLeaveType.SelectedValue = 14 Then 'ecq - quarantine
                                            .IsFitToWork = False
                                        Else
                                            If isUnfitToWork = True Then
                                                .IsFitToWork = False
                                            Else
                                                .IsFitToWork = True
                                            End If
                                        End If
                                    End With

                                    Dim newRowLeaveFiling As LeaveFilingRow = Me.dsLeaveFiling.LeaveFiling.NewLeaveFilingRow

                                    With newRowLeaveFiling
                                        .DateCreated = dbLeaveFiling.GetServerDate
                                        .ScreenId = screenId
                                        .EmployeeId = employeeId
                                        .DepartmentId = departmentId
                                        .TeamId = teamId
                                        .StartDate = CDate(txtAbsentFrom.Text)
                                        .EndDate = CDate(txtAbsentTo.Text)
                                        .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        .Reason = txtReason.Text.Trim
                                        .LeaveCredits = GetLeaveCredits(employeeId)
                                        .LeaveBalance = GetLeaveBalance(employeeId)
                                        .ClinicIsApproved = 1
                                        .ClinicId = screenBy
                                        .ClinicApprovalDate = dbLeaveFiling.GetServerDate
                                        .ClinicRemarks = txtDiagnosis.Text.Trim
                                        .IsLateFiling = 1
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .SetModifiedByNull()
                                        .SetModifiedDateNull()
                                        .IsEncoded = 0
                                        .IsDone = 0

                                        .SuperiorIsApproved1 = 0
                                        .SetSuperiorApprovalDate1Null()
                                        .SetSuperiorRemarks1Null()

                                        .SuperiorIsApproved2 = 0
                                        .SetSuperiorApprovalDate2Null()
                                        .SetSuperiorRemarks2Null()

                                        .ManagerIsApproved = 0
                                        .SetManagerApprovalDateNull()
                                        .SetManagerRemarksNull()

                                        Dim prmCountRecipient(2) As SqlParameter
                                        prmCountRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                        prmCountRecipient(0).Value = departmentId
                                        prmCountRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                        prmCountRecipient(1).Value = teamId
                                        prmCountRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                        prmCountRecipient(2).Value = positionId

                                        Dim count As Integer = dbLeaveFiling.ExecuteScalar("SELECT COUNT(RecipientId) AS Count FROM dbo.Recipient " &
                                                                                           "WHERE DepartmentId = @DepartmentId AND TeamId = @TeamId AND PositionId = @PositionId ",
                                                                                           CommandType.Text, prmCountRecipient)

                                        If count = 0 Then 'inform the developer regarding automatic filing with no recipients
                                            frmScreenList.SendDevNotif(employeeId, txtEmployeeName.Text.ToString.Trim, cmbLeaveType.SelectedValue,
                                                                       cmbLeaveType.Text, departmentId, departmentName, teamId, teamName, positionId, positionName)
                                            'set the last approver based on the majority of other records per department
                                            Dim managerId As Integer = 0
                                            Dim prmDeptId(0) As SqlParameter
                                            prmDeptId(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmDeptId(0).Value = departmentId

                                            managerId = dbLeaveFiling.ExecuteScalar("SELECT TOP 1 ManagerId FROM dbo.Recipient " &
                                                                                    "WHERE DepartmentId = @DepartmentId GROUP BY ManagerId " &
                                                                                    "ORDER BY COUNT(RecipientId) DESC",
                                                                                    CommandType.Text, prmDeptId)

                                            If employeeId = managerId Then 'employee is a manager, set dgm as the approver
                                                .ManagerId = 70
                                            Else
                                                .ManagerId = managerId
                                            End If

                                            .RoutingStatusId = 3
                                        Else
                                            Dim prmRecipient(2) As SqlParameter
                                            prmRecipient(0) = New SqlParameter("@DepartmentId", SqlDbType.Int)
                                            prmRecipient(0).Value = departmentId
                                            prmRecipient(1) = New SqlParameter("@TeamId", SqlDbType.Int)
                                            prmRecipient(1).Value = teamId
                                            prmRecipient(2) = New SqlParameter("PositionId", SqlDbType.Int)
                                            prmRecipient(2).Value = positionId

                                            Using readerRecipient As IDataReader = dbLeaveFiling.ExecuteReader("RdRecipient", CommandType.StoredProcedure, prmRecipient)
                                                Dim superiorId1 As Integer = 0
                                                Dim superiorId2 As Integer = 0
                                                Dim managerId As Integer = 0

                                                While readerRecipient.Read
                                                    If readerRecipient.Item("SuperiorId1") Is DBNull.Value Then 'no superior 1
                                                        .SetSuperiorId1Null()

                                                        If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()

                                                            If employeeId = readerRecipient.Item("ManagerId") Then 'employee is a manager, set dgm as the approver
                                                                .RoutingStatusId = 3
                                                                .ManagerId = 70 'dgm
                                                            Else
                                                                .RoutingStatusId = 3
                                                                .ManagerId = readerRecipient.Item("ManagerId")
                                                            End If
                                                        Else
                                                            If employeeId = readerRecipient.Item("SuperiorId2") Then
                                                                .RoutingStatusId = 3
                                                                .SetSuperiorId2Null()
                                                            Else
                                                                .RoutingStatusId = 4
                                                                .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                            End If
                                                        End If
                                                    Else 'with superior 1
                                                        If employeeId = readerRecipient.Item("SuperiorId1") Then
                                                            .RoutingStatusId = 4
                                                            .SetSuperiorId1Null()
                                                        Else
                                                            .RoutingStatusId = 5
                                                            .SuperiorId1 = readerRecipient.Item("SuperiorId1")
                                                        End If

                                                        If readerRecipient.Item("SuperiorId2") Is DBNull.Value Then
                                                            .SetSuperiorId2Null()
                                                        Else
                                                            .SuperiorId2 = readerRecipient.Item("SuperiorId2")
                                                        End If
                                                    End If

                                                    .ManagerId = readerRecipient.Item("ManagerId")
                                                    managerId = readerRecipient.Item("ManagerId")
                                                End While
                                                readerRecipient.Close()
                                            End Using
                                        End If
                                    End With
                                    Me.dsLeaveFiling.LeaveFiling.AddLeaveFilingRow(newRowLeaveFiling)
                                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)

                                    If newRowLeaveFiling.RoutingStatusId = 3 Then
                                        If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                            frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                            newRowLeaveFiling.ManagerId,
                                                                            cmbLeaveType.Text,
                                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                            departmentName,
                                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                            txtReason.Text.Trim)
                                        Else
                                            frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                            newRowLeaveFiling.ManagerId,
                                                                            cmbLeaveType.Text,
                                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                            departmentName,
                                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                            CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                            txtReason.Text.Trim)
                                        End If
                                    ElseIf newRowLeaveFiling.RoutingStatusId = 4 Then
                                        If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                            frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                            newRowLeaveFiling.SuperiorId2,
                                                                            cmbLeaveType.Text,
                                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                            departmentName,
                                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                            txtReason.Text.Trim)
                                        Else
                                            frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                            newRowLeaveFiling.SuperiorId2,
                                                                            cmbLeaveType.Text,
                                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                            departmentName,
                                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                            CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                            txtReason.Text.Trim)
                                        End If
                                    ElseIf newRowLeaveFiling.RoutingStatusId = 5 Then
                                        If CDate(txtAbsentFrom.Text).Date.Equals(CDate(txtAbsentTo.Text).Date) Then
                                            frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                            newRowLeaveFiling.SuperiorId1,
                                                                            cmbLeaveType.Text,
                                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                            departmentName,
                                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy"),
                                                                            txtReason.Text.Trim)
                                        Else
                                            frmScreenList.SendApproverNotif(newRowLeaveFiling.LeaveFileId,
                                                                            newRowLeaveFiling.SuperiorId1,
                                                                            cmbLeaveType.Text,
                                                                            StrConv(txtEmployeeName.Text.Trim, VbStrConv.ProperCase),
                                                                            departmentName,
                                                                            CDate(txtAbsentFrom.Text).Date.ToString("MMMM dd, yyyy") & " - " &
                                                                            CDate(txtAbsentTo.Text).Date.ToString("MMMM dd, yyyy"),
                                                                            txtReason.Text.Trim)
                                        End If
                                    End If
                                Else 'other leave types
                                    With rowScreening
                                        .ModifiedBy = screenBy
                                        .ModifiedDate = dbLeaveFiling.GetServerDate
                                        .EmployeeId = employeeId
                                        .AbsentFrom = CDate(txtAbsentFrom.Text)
                                        .AbsentTo = CDate(txtAbsentTo.Text)
                                        .LeaveTypeId = cmbLeaveType.SelectedValue
                                        .Reason = txtReason.Text.Trim
                                        .Diagnosis = txtDiagnosis.Text.Trim

                                        Select Case cmbLeaveType.SelectedValue
                                            Case 12, 15, 16
                                                .Quantity = 0.5
                                            Case Else
                                                .Quantity = GetTotalDays(CDate(txtAbsentFrom.Text), CDate(txtAbsentTo.Text))
                                        End Select

                                        If isUnfitToWork = True Then
                                            .IsFitToWork = False
                                        Else
                                            .IsFitToWork = True
                                        End If
                                    End With
                                End If
                            End If
                        End If
                    End If
                End If

                Me.Validate()
                Me.bsScreening.EndEdit()
                Me.bsLeaveFiling.EndEdit()

                If Me.dsLeaveFiling.HasChanges Then
                    Me.adpScreening.Update(Me.dsLeaveFiling.Screening)
                    Me.adpLeaveFiling.Update(Me.dsLeaveFiling.LeaveFiling)
                    Me.dsLeaveFiling.AcceptChanges()
                    Me.DialogResult = DialogResult.OK
                End If
            End If

            employeeId = 0
            departmentId = 0
            teamId = 0
            positionId = 0
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Functions"

    'set the absent date to the last working date - excluding sunday, company holidays and legal holidays
    Private Function GetLastWorkingDay(subjectDate As DateTime) As Date
        Try
            subjectDate = subjectDate.AddDays(-1)
            While IsHoliday(subjectDate) Or IsWeekend(subjectDate)
                subjectDate = subjectDate.AddDays(-1)
            End While
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return subjectDate
    End Function

    'get leave balance
    Private Function GetLeaveBalance(employeeId As Integer) As Integer
        Dim leaveBalance As Double = 0

        Try
            If Not cmbLeaveType.SelectedValue = 0 Then
                'Dim prmBalance(3) As SqlParameter
                'prmBalance(0) = New SqlParameter("@CompanyId", SqlDbType.Int)
                'prmBalance(0).Value = 1
                'prmBalance(1) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                'prmBalance(1).Value = employeeId
                'prmBalance(2) = New SqlParameter("@LeaveTypeId", SqlDbType.Int)
                'prmBalance(2).Value = cmbLeaveType.SelectedValue
                'prmBalance(3) = New SqlParameter("@Date", SqlDbType.Date)
                'prmBalance(3).Value = DBNull.Value

                'leaveBalance = dbJeonsoft.ExecuteFunction(Of Double)("dbo.fnGetLeaveBalance", prmBalance)

                Dim prmBalance(2) As SqlParameter
                prmBalance(0) = New SqlParameter("@CompanyId", SqlDbType.Int)
                prmBalance(0).Value = 1
                prmBalance(1) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                prmBalance(1).Value = employeeId
                prmBalance(2) = New SqlParameter("@LeaveTypeId", SqlDbType.Int)
                prmBalance(2).Value = cmbLeaveType.SelectedValue

                leaveBalance = dbJeonsoft.ExecuteScalar("SELECT Balance FROM dbo.tblLeaveBalances WHERE EmployeeId = @EmployeeId AND LeaveTypeId = @LeaveTypeId " &
                                                        "AND CompanyId = @CompanyId", CommandType.Text, prmBalance)
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return leaveBalance
    End Function

    'get leave credits
    Private Function GetLeaveCredits(employeeId As Integer) As Integer
        Dim leaveCredits As Double = 0

        Try
            If Not cmbLeaveType.SelectedValue = 0 Then
                Dim prmCredits(2) As SqlParameter
                prmCredits(0) = New SqlParameter("@EmployeeId", SqlDbType.Int)
                prmCredits(0).Value = employeeId
                prmCredits(1) = New SqlParameter("@LeaveTypeId", SqlDbType.Int)
                prmCredits(1).Value = cmbLeaveType.SelectedValue
                prmCredits(2) = New SqlParameter("@Year", SqlDbType.Int)
                prmCredits(2).Value = Year(dbLeaveFiling.GetServerDate)

                leaveCredits = dbJeonsoft.ExecuteScalar("SELECT TOP 1 EndBalance FROM dbo.tblLeaveLedger WHERE YEAR(Date) = YEAR(GETDATE()) AND " &
                                                        "EmployeeId = @EmployeeId AND LeaveTypeId = @LeaveTypeId ORDER BY Date ASC",
                                                        CommandType.Text, prmCredits)
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return leaveCredits
    End Function

    'get the total number of days from start date up to end date - excluding holidays and sundays
    Private Function GetTotalDays(startDate As Date, endDate As Date) As Integer
        Dim countDays As Integer = 0

        Try
            If startDate.Date.Equals(endDate.Date) Then
                countDays = 1
            Else
                For i As Integer = 0 To (endDate - startDate).Days
                    If Not IsHoliday(startDate) Then
                        If Not IsWeekend(startDate) Then
                            countDays += 1
                        End If
                    End If
                    startDate = startDate.AddDays(1)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        txtTotalDays.Text = countDays

        Return countDays
    End Function

    Private Function IsHoliday(subjectDate As Date) As Boolean
        Dim count As Integer

        Try
            Dim prmDate(0) As SqlParameter
            prmDate(0) = New SqlParameter("@HolidayDate", SqlDbType.Date)
            prmDate(0).Value = subjectDate.ToShortDateString
            count = 0
            count = dbLeaveFiling.ExecuteScalar("SELECT COUNT(HolidayId) FROM dbo.Holiday WHERE HolidayDate = @HolidayDate", CommandType.Text, prmDate)
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsWeekend(subjectDate As Date) As Boolean
        If subjectDate.DayOfWeek.Equals(DayOfWeek.Sunday) Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

#Region "UI"

    Private Sub chkNotFtw_Enter(sender As Object, e As EventArgs) Handles chkNotFtw.Enter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub chkNotFtw_Leave(sender As Object, e As EventArgs) Handles chkNotFtw.Leave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub chkNotFtw_MouseEnter(sender As Object, e As EventArgs) Handles chkNotFtw.MouseEnter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub chkNotFtw_MouseLeave(sender As Object, e As EventArgs) Handles chkNotFtw.MouseLeave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbLeaveType_Enter(sender As Object, e As EventArgs) Handles cmbLeaveType.Enter
        lblLeaveType.ForeColor = Color.White
        lblLeaveType.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbLeaveType_Leave(sender As Object, e As EventArgs) Handles cmbLeaveType.Leave
        lblLeaveType.ForeColor = Color.Black
        lblLeaveType.BackColor = SystemColors.Control
    End Sub

    Private Sub lblNotFtw_Enter(sender As Object, e As EventArgs) Handles lblNotFtw.Enter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub lblNotFtw_Leave(sender As Object, e As EventArgs) Handles lblNotFtw.Leave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub lblNotFtw_MouseEnter(sender As Object, e As EventArgs) Handles lblNotFtw.MouseEnter
        lblNotFtw.ForeColor = Color.White
        lblNotFtw.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub lblNotFtw_MouseLeave(sender As Object, e As EventArgs) Handles lblNotFtw.MouseLeave
        lblNotFtw.ForeColor = Color.Black
        lblNotFtw.BackColor = SystemColors.Control
    End Sub

    Private Sub txtAbsentFrom_Enter(sender As Object, e As EventArgs) Handles txtAbsentFrom.Enter
        lblAbsentFrom.ForeColor = Color.White
        lblAbsentFrom.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtAbsentFrom_Leave(sender As Object, e As EventArgs) Handles txtAbsentFrom.Leave
        lblAbsentFrom.ForeColor = Color.Black
        lblAbsentFrom.BackColor = SystemColors.Control
    End Sub

    Private Sub txtAbsentTo_Enter(sender As Object, e As EventArgs) Handles txtAbsentTo.Enter
        lblAbsentTo.ForeColor = Color.White
        lblAbsentTo.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtAbsentTo_Leave(sender As Object, e As EventArgs) Handles txtAbsentTo.Leave
        lblAbsentTo.ForeColor = Color.Black
        lblAbsentTo.BackColor = SystemColors.Control
    End Sub

    Private Sub txtDiagnosis_Enter(sender As Object, e As EventArgs) Handles txtDiagnosis.Enter
        lblDiagnosis.ForeColor = Color.White
        lblDiagnosis.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtDiagnosis_Leave(sender As Object, e As EventArgs) Handles txtDiagnosis.Leave
        lblDiagnosis.ForeColor = Color.Black
        lblDiagnosis.BackColor = SystemColors.Control
    End Sub

    Private Sub txtEmployeeName_Enter(sender As Object, e As EventArgs) Handles txtEmployeeName.Enter
        lblEmployeeName.ForeColor = Color.White
        lblEmployeeName.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtEmployeeName_Leave(sender As Object, e As EventArgs) Handles txtEmployeeName.Leave
        lblEmployeeName.ForeColor = Color.Black
        lblEmployeeName.BackColor = SystemColors.Control
    End Sub

    Private Sub txtEmployeeScanId_Enter(sender As Object, e As EventArgs) Handles txtEmployeeScanId.Enter
        lblEmployeeScanId.ForeColor = Color.White
        lblEmployeeScanId.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtEmployeeScanId_Leave(sender As Object, e As EventArgs) Handles txtEmployeeScanId.Leave
        lblEmployeeScanId.ForeColor = Color.Black
        lblEmployeeScanId.BackColor = SystemColors.Control
    End Sub

    Private Sub txtReason_Enter(sender As Object, e As EventArgs) Handles txtReason.Enter
        lblReason.ForeColor = Color.White
        lblReason.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub txtReason_Leave(sender As Object, e As EventArgs) Handles txtReason.Leave
        lblReason.ForeColor = Color.Black
        lblReason.BackColor = SystemColors.Control
    End Sub

#End Region

End Class