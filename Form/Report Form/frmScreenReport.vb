Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms
Imports BlackCoffeeLibrary

Public Class frmScreenReport
    Private connection As New clsConnection
    Private dbLeaveFiling As New SqlDbMethod(connection.ClientConnection)
    Private dbJeonsoft As New SqlDbMethod(connection.JeonsoftConnection)
    Private dbMain As New Main

    Private serverDate As DateTime = dbLeaveFiling.GetServerDate
    Private dsScreeningReport As New dsScreeningReport
    Private bsScreeningReport As New BindingSource

    'report paramaters
    Private periodCovered As String = String.Empty
    Private leaveType As String = String.Empty
    Private employmentType As String = String.Empty
    Private employeeName As String = String.Empty
    Private status As String = String.Empty

    'dictionaries
    Private dicDateType As New Dictionary(Of String, Integer)
    Private dicEmploymentType As New Dictionary(Of String, Integer)
    Private dicStatus As New Dictionary(Of String, Integer)

    Private Sub frmHealthScreeningReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            AddHandler Me.SizeChanged, AddressOf Me_SizeChanged
            Me.MaximizeBox = False

            dicDateType.Add(" Absent Date", 1)
            dicDateType.Add(" Screen Date", 2)
            dicDateType.Add(" Medical Cert Date", 3)
            cmbDateType.DisplayMember = "Key"
            cmbDateType.ValueMember = "Value"
            cmbDateType.DataSource = New BindingSource(dicDateType, Nothing)

            dbLeaveFiling.FillCmbWithCaption("RdLeaveType", CommandType.StoredProcedure, "LeaveTypeId", "LeaveTypeName", cmbLeaveType, "< Select Leave Type >")

            dbJeonsoft.FillCmbWithCaption("SELECT Id, (TRIM(EmployeeCode) + '  ' + (FirstName + ' ' + ISNULL(SUBSTRING(CASE WHEN LEN(TRIM(MiddleName)) = 0 THEN NULL " &
                                          "WHEN TRIM(MiddleName) = '-' THEN NULL ELSE TRIM(MiddleName) END, 1, 1) + '. ' , '') + LastName)) AS Name FROM dbo.tblEmployees WHERE " &
                                          "Active = 1 And EmployeeCode Is NOT NULL", CommandType.Text, "Id", "Name", cmbEmployeeName, "")

            dicEmploymentType.Add("< Select Employment > ", 0)
            dicEmploymentType.Add("Direct", 1)
            dicEmploymentType.Add("Agency", 2)
            cmbEmploymentType.DisplayMember = "Key"
            cmbEmploymentType.ValueMember = "Value"
            cmbEmploymentType.DataSource = New BindingSource(dicEmploymentType, Nothing)

            dicStatus.Add("< Select Status > ", 0)
            dicStatus.Add("Fit To Work", 1)
            dicStatus.Add("Not Fit To Work", 2)
            cmbStatus.DisplayMember = "Key"
            cmbStatus.ValueMember = "Value"
            cmbStatus.DataSource = New BindingSource(dicStatus, Nothing)

            rptViewer.LocalReport.ReportPath = ""

            Me.ActiveControl = btnGenerate
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHealthScreeningReport_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode.Equals(Keys.F10) Then
            e.Handled = True
            btnGenerate.PerformClick()
        End If
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Try
            If dtpStartDate.Value.Date > dtpEndDate.Value.Date Then
                MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf dtpStartDate.Value.Date = dtpEndDate.Value.Date Then
                GenerateReport()
            Else
                GenerateReport()
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GenerateReport()
        Try
            Dim prmRpt(6) As SqlParameter
            prmRpt(0) = New SqlParameter("@AbsentFrom", SqlDbType.Date)
            prmRpt(0).Value = dtpStartDate.Value.Date
            prmRpt(1) = New SqlParameter("@AbsentTo", SqlDbType.Date)
            prmRpt(1).Value = dtpEndDate.Value.Date
            prmRpt(2) = New SqlParameter("@DateType", SqlDbType.Char)
            prmRpt(2).Value = cmbDateType.SelectedValue
            prmRpt(3) = New SqlParameter("@EmployeeId", SqlDbType.Int)
            prmRpt(3).Value = GetEmployee()
            prmRpt(4) = New SqlParameter("@LeaveTypeId", SqlDbType.Int)
            prmRpt(4).Value = GetLeaveType()
            prmRpt(5) = New SqlParameter("@Status", SqlDbType.Int)
            prmRpt(5).Value = GetStatus()
            prmRpt(6) = New SqlParameter("@EmploymentTypeId", SqlDbType.Int)
            prmRpt(6).Value = GetEmploymentType()

            Dim dtReport As New DataTable
            dtReport = dbLeaveFiling.FillDataTable("RptScreening", CommandType.StoredProcedure, prmRpt)

            If dtReport.Rows.Count = 0 Then
                MessageBox.Show("No records found.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            rptViewer.ProcessingMode = ProcessingMode.Local
            rptViewer.LocalReport.ReportPath = "ReportFile\Screening.rdlc"
            rptViewer.LocalReport.DataSources.Clear()
            rptViewer.LocalReport.DataSources.Add(New ReportDataSource("RptScreening", dtReport))

            Dim rptParam As New ReportParameterCollection
            Dim periodCovered As String = String.Empty
            Dim leaveType As String = String.Empty
            Dim employmentType As String = String.Empty

            If dtpStartDate.Value.Date.Equals(dtpEndDate.Value.Date) Then
                periodCovered = dtpStartDate.Value.ToString("MMMM dd, yyyy")
            Else
                periodCovered = dtpStartDate.Value.ToString("MMMM dd, yyyy") & " to " & dtpEndDate.Value.ToString("MMMM dd, yyyy")
            End If

            If Not cmbLeaveType.SelectedValue = 0 Then
                leaveType = cmbLeaveType.Text
            Else
                leaveType = " "
            End If

            If Not cmbEmploymentType.SelectedValue = 0 Then
                If cmbEmploymentType.SelectedValue = 1 Then
                    employmentType = "Direct"
                ElseIf cmbEmploymentType.SelectedValue = 2 Then
                    employmentType = "Agency"
                End If
            Else
                employmentType = " "
            End If

            If Not cmbEmployeeName.SelectedValue = 0 Then
                employeeName = cmbEmployeeName.Text
            Else
                employeeName = " "
            End If

            If Not cmbStatus.SelectedValue = 0 Then
                status = cmbStatus.Text
            Else
                status = " "
            End If

            rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("PeriodCovered", periodCovered))
            rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("LeaveType", leaveType))
            rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("EmploymentType", employmentType))
            rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("EmployeeName", employeeName))
            rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("Status", status))
            rptViewer.LocalReport.SetParameters(rptParam)

            rptViewer.SetDisplayMode(DisplayMode.PrintLayout)
            rptViewer.ZoomMode = ZoomMode.PageWidth
            rptViewer.LocalReport.DisplayName = "Monitoring Report"
            rptViewer.RefreshReport()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetEmployee() As Object
        If cmbEmployeeName.SelectedValue = 0 Then Return Nothing Else Return cmbEmployeeName.SelectedValue
    End Function

    Private Function GetLeaveType() As Object
        If cmbLeaveType.SelectedValue = 0 Then Return Nothing Else Return cmbLeaveType.SelectedValue
    End Function

    Private Function GetStatus() As Object
        If cmbStatus.SelectedValue = 0 Then Return Nothing Else Return cmbStatus.SelectedValue
    End Function

    Private Function GetEmploymentType() As Object
        If cmbEmploymentType.SelectedValue = 0 Then Return Nothing Else Return cmbEmploymentType.SelectedValue
    End Function

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        cmbDateType.SelectedValue = 1
        dtpStartDate.Value = Date.Now.Date
        dtpEndDate.Value = Date.Now.Date
        cmbLeaveType.SelectedValue = 0
        cmbEmploymentType.SelectedValue = 0
        cmbStatus.SelectedValue = 0
        cmbEmployeeName.SelectedValue = 0
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    'set the window state to maximized without overlapping the taskbar
    Private Sub Me_SizeChanged(sender As Object, e As EventArgs)
        If Me.WindowState = FormWindowState.Minimized Then
            Me.MaximizeBox = True

        ElseIf Me.WindowState = FormWindowState.Maximized Then
            Me.MaximizeBox = False
        End If
    End Sub

    Private Sub cmbDateType_Enter(sender As Object, e As EventArgs) Handles cmbDateType.Enter
        lblDate.ForeColor = Color.White
        lblDate.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbDateType_Leave(sender As Object, e As EventArgs) Handles cmbDateType.Leave
        lblDate.ForeColor = Color.Black
        lblDate.BackColor = SystemColors.Control
    End Sub

    Private Sub dtpStartDate_Enter(sender As Object, e As EventArgs) Handles dtpStartDate.Enter
        lblStartDate.ForeColor = Color.White
        lblStartDate.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub dtpStartDate_Leave(sender As Object, e As EventArgs) Handles dtpStartDate.Leave
        lblStartDate.ForeColor = Color.Black
        lblStartDate.BackColor = SystemColors.Control
    End Sub

    Private Sub dtpEndDate_Enter(sender As Object, e As EventArgs) Handles dtpEndDate.Enter
        lblEndDate.ForeColor = Color.White
        lblEndDate.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub dtpEndDate_Leave(sender As Object, e As EventArgs) Handles dtpEndDate.Leave
        lblEndDate.ForeColor = Color.Black
        lblEndDate.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbLeaveType_Enter(sender As Object, e As EventArgs) Handles cmbLeaveType.Enter
        lblLeaveType.ForeColor = Color.White
        lblLeaveType.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbLeaveType_Leave(sender As Object, e As EventArgs) Handles cmbLeaveType.Leave
        lblLeaveType.ForeColor = Color.Black
        lblLeaveType.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbEmploymentType_Enter(sender As Object, e As EventArgs) Handles cmbEmploymentType.Enter
        lblEmploymentType.ForeColor = Color.White
        lblEmploymentType.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbEmploymentType_Leave(sender As Object, e As EventArgs) Handles cmbEmploymentType.Leave
        lblEmploymentType.ForeColor = Color.Black
        lblEmploymentType.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbStatus_Enter(sender As Object, e As EventArgs) Handles cmbStatus.Enter
        lblStatus.ForeColor = Color.White
        lblStatus.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbStatus_Leave(sender As Object, e As EventArgs) Handles cmbStatus.Leave
        lblStatus.ForeColor = Color.Black
        lblStatus.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbEmployeeName_Enter(sender As Object, e As EventArgs) Handles cmbEmployeeName.Enter
        lblEmployeeName.ForeColor = Color.White
        lblEmployeeName.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub cmbEmployeeName_Leave(sender As Object, e As EventArgs) Handles cmbEmployeeName.Leave
        lblEmployeeName.ForeColor = Color.Black
        lblEmployeeName.BackColor = SystemColors.Control
    End Sub

    Private Sub cmbLeaveType_Validated(sender As Object, e As EventArgs) Handles cmbLeaveType.Validated
        If cmbLeaveType.SelectedValue = 0 Then
            cmbLeaveType.SelectedValue = 0
        End If
    End Sub

    Private Sub cmbEmploymentType_Validated(sender As Object, e As EventArgs) Handles cmbEmploymentType.Validated
        If cmbEmploymentType.SelectedValue = 0 Then
            cmbEmploymentType.SelectedValue = 0
        End If
    End Sub

    Private Sub cmbStatus_Validated(sender As Object, e As EventArgs) Handles cmbStatus.Validated
        If cmbStatus.SelectedValue = 0 Then
            cmbStatus.SelectedValue = 0
        End If
    End Sub

    Private Sub cmbEmployeeName_Validated(sender As Object, e As EventArgs) Handles cmbEmployeeName.Validated
        If cmbEmployeeName.SelectedValue = 0 Then
            cmbEmployeeName.SelectedValue = 0
        End If
    End Sub

End Class