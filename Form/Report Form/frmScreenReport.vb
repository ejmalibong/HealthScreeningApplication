Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms
Imports BlackCoffeeLibrary
Imports SickLeaveScreening
Imports SickLeaveScreening.dsScreeningReport
Imports SickLeaveScreening.dsScreeningReportTableAdapters

Public Class frmScreenReport
    Private connection As New clsConnection
    Private dbLeaveFiling As New SqlDbMethod(connection.LocalConnection)
    Private dbMain As New Main

    Private serverDate As DateTime = dbLeaveFiling.GetServerDate

    Private dsScreeningReport As New dsScreeningReport
    Private adpScreeningReport As New VwScreeningTableAdapter
    Private dtScreeningReport As New VwScreeningDataTable
    Private bsScreeningReport As New BindingSource

    'report paramaters
    Private query As String = String.Empty
    Private periodCovered As String = String.Empty
    Private leaveType As String = String.Empty
    Private employeeType As String = String.Empty

    'dictionaries
    Private dicEmploymentType As New Dictionary(Of String, Integer)

    Private Sub frmHealthScreeningReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            dbLeaveFiling.FillCmbWithCaption("RdLeaveType", CommandType.StoredProcedure, "LeaveTypeId", "LeaveTypeName", cmbLeaveType, "< Select Leave Type >")

            dicEmploymentType.Add("< Select Employment > ", 0)
            dicEmploymentType.Add("Direct", 1)
            dicEmploymentType.Add("Agency", 2)
            cmbEmploymentType.DisplayMember = "Key"
            cmbEmploymentType.ValueMember = "Value"
            cmbEmploymentType.DataSource = New BindingSource(dicEmploymentType, Nothing)

            Me.adpScreeningReport.Fill(Me.dsScreeningReport.VwScreening)
            Me.bsScreeningReport.DataSource = Me.dsScreeningReport
            Me.bsScreeningReport.DataMember = dtScreeningReport.TableName

            Me.ActiveControl = btnGenerate
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHealthScreeningReport_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode.Equals(Keys.Enter) Then
            e.Handled = True
            Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
        ElseIf e.KeyCode.Equals(Keys.F10) Then
            e.Handled = True
            btnGenerate.PerformClick()
        End If
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Try
            If dtpStartDate.Value.Date > dtpEndDate.Value.Date Then
                MessageBox.Show("Start date is later than end date.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf dtpStartDate.Value.Date = dtpEndDate.Value.Date Then
                GoTo GenerateReport
            Else

GenerateReport:
                query = "ScreenDate >= '" + dbMain.FormatDate(dtpStartDate.Value.Date, True) + "' AND ScreenDate < '" + dbMain.FormatDate(dtpEndDate.Value.Date, False) + "'"

                If dtpStartDate.Value.Date.Equals(dtpEndDate.Value.Date) Then
                    periodCovered = dtpStartDate.Value.ToString("MMMM dd, yyyy")
                Else
                    periodCovered = dtpStartDate.Value.ToString("MMMM dd, yyyy") & " to " & dtpEndDate.Value.ToString("MMMM dd, yyyy")
                End If

                If Not cmbLeaveType.SelectedValue = 0 Then
                    query += " AND LeaveTypeId = '" & cmbLeaveType.SelectedValue & "'"
                    leaveType = cmbLeaveType.Text
                Else
                    leaveType = " "
                End If

                If Not cmbEmploymentType.SelectedValue = 0 Then
                    If cmbEmploymentType.SelectedValue = 1 Then
                        query += " AND EmployeeId <> 0"
                        employeeType = "Direct"
                    ElseIf cmbEmploymentType.SelectedValue = 2 Then
                        query += " AND EmployeeId = 0"
                        employeeType = "Agency"
                    End If
                Else
                    employeeType = " "
                End If

                If chkNotFtw.CheckState = CheckState.Checked Then
                    query += " AND IsFitToWork = 0"
                End If

                Me.bsScreeningReport.Filter = String.Format(query)
                Me.bsScreeningReport.Sort = "ScreenDate ASC"

                If Me.bsScreeningReport.Count > 0 Then
                    rptViewer.LocalReport.ReportPath = "ReportFile\Screening.rdlc"
                    rptViewer.LocalReport.DataSources.Clear()
                    rptViewer.LocalReport.DataSources.Add(New ReportDataSource(dtScreeningReport.TableName, Me.bsScreeningReport))

                    Dim _rptParam As New ReportParameterCollection
                    _rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("PeriodCovered", periodCovered))
                    _rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("LeaveType", leaveType))
                    _rptParam.Add(New Microsoft.Reporting.WinForms.ReportParameter("EmploymentType", employeeType))
                    rptViewer.LocalReport.SetParameters(_rptParam)

                    rptViewer.SetDisplayMode(DisplayMode.PrintLayout)
                    rptViewer.ZoomMode = ZoomMode.PageWidth
                    rptViewer.LocalReport.DisplayName = "Monitoring Report"
                    rptViewer.RefreshReport()
                Else
                    MessageBox.Show("No records found.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnRemoveFilters.Click
        dbMain.FormReset(Me)
        chkNotFtw.CheckState = CheckState.Unchecked
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

End Class