Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports BlackCoffeeLibrary
Imports SickLeaveScreening
Imports SickLeaveScreening.dsLeaveFiling
Imports SickLeaveScreening.dsLeaveFilingTableAdapters

Public Class frmDoctor
    Private connection As New clsConnection
    Private dbLeaveFiling As New SqlDbMethod(connection.ServerConnection)
    Private dbMain As New Main

    Private dsLeaveFiling As New dsLeaveFiling
    Private adpClinic As New ClinicTableAdapter
    Private dtClinic As New ClinicDataTable
    Private WithEvents bsClinic As New BindingSource

    Private pageSize As Integer
    Private pageIndex As Integer
    Private totalCount As Integer
    Private pageCount As Integer
    Private indexScroll As Integer = 0
    Private indexPosition As Integer = 0

    Private dictionary As New Dictionary(Of String, Integer)

    Private isFilterByName As Boolean = False

    Private isExists As Boolean = True
    Private isEnableValidation As Boolean = True

    Private Sub frmDoctor_Load(sender As Object, e As EventArgs) Handles Me.Load
        pageIndex = 0
        pageSize = 100
        BindPage()

        dbMain.EnableDoubleBuffered(dgvList)

        Me.dgvList.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Me.dgvList.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        Me.ActiveControl = dgvList
    End Sub

    Private Sub frmDoctor_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode.Equals(Keys.F2) Then
            e.Handled = True
            btnAdd.PerformClick()
        ElseIf e.KeyCode.Equals(Keys.F8) Then
            e.Handled = True
            btnDelete.PerformClick()
        ElseIf e.KeyCode.Equals(Keys.F10) Then
            e.Handled = True
            btnSave.PerformClick()
        End If
    End Sub

    Private Sub dgvList_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvList.CellValidating
        Try
            If isEnableValidation = True Then
                If e.ColumnIndex = 1 Then 'employee code
                    If String.IsNullOrEmpty(e.FormattedValue.ToString.Trim) Then
                        MessageBox.Show("Badge number is required.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        e.Cancel = True
                    End If

                    If dgvList.IsCurrentCellDirty Then
                        dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
                    End If

                ElseIf e.ColumnIndex = 2 Then 'employee name
                    If String.IsNullOrEmpty(e.FormattedValue.ToString.Trim) Then
                        MessageBox.Show("Name is required.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        e.Cancel = True
                    End If

                    If dgvList.IsCurrentCellDirty Then
                        dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
                    End If

                ElseIf e.ColumnIndex = 4 Then 'password
                    If dgvList.IsCurrentCellDirty Then
                        dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
                    End If

                    isExists = False

                    For x As Integer = 0 To dgvList.Rows.Count - 1
                        For y As Integer = 0 To dgvList.Rows.Count - 1
                            If y <> x AndAlso dgvList.Rows(x).Cells(4).Value.ToString.ToLower = dgvList.Rows(y).Cells(4).Value.ToString.ToLower Then
                                isExists = True
                            End If
                        Next
                    Next

                    If isExists = True Then
                        MessageBox.Show("Password already used.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        e.Cancel = True
                    End If

                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvList_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvList.DataError
        e.Cancel = False
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Try
            MessageBox.Show("Under maintenance.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub

            'Me.bsClinic.AddNew()
            'Me.bsClinic.MoveLast()
            'dgvList.BeginEdit(True)
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Me.Validate()
            Me.bsClinic.EndEdit()
            If Me.dsLeaveFiling.HasChanges Then
                Me.adpClinic.Update(dsLeaveFiling.Clinic)
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Try
            Me.Validate()

            If dgvList.NewRowIndex Then
                dgvList.CancelEdit()
                Me.bsClinic.CancelEdit()
            End If

            If dsLeaveFiling.HasChanges Then
                Dim result As DialogResult = MessageBox.Show("Discard changes?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)

                If result = Windows.Forms.DialogResult.Yes Then
                    dgvList.CancelEdit()
                    Me.bsClinic.CancelEdit()
                    Me.dsLeaveFiling.RejectChanges()

                ElseIf result = Windows.Forms.DialogResult.No Then
                    dgvList.CurrentRow.Cells(1).Selected = True
                    dgvList.BeginEdit(True)
                    Return
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If Me.bsClinic.Current Is Nothing Then
                Exit Sub
            End If

            Dim currentRow = CType(Me.bsClinic.Current, DataRowView).Row
            Dim state = currentRow.RowState

            Select Case state
                Case DataRowState.Added
                    Me.bsClinic.RemoveCurrent()
                Case DataRowState.Deleted
                    MessageBox.Show("Item is already deleted.", "")
                Case DataRowState.Detached
                    Me.bsClinic.CancelEdit()
                Case DataRowState.Modified, DataRowState.Unchanged
                    If dgvList.SelectedCells.Count > 0 AndAlso dgvList.SelectedCells(0).RowIndex = dgvList.NewRowIndex Then
                        Me.bsClinic.CancelEdit()
                        Exit Sub
                    End If

                    Dim message = String.Format("Delete {0}?", bsClinic.Current("EmployeeName"))
                    If MessageBox.Show(message, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        Me.bsClinic.RemoveCurrent()
                        Me.adpClinic.Update(Me.dsLeaveFiling.Clinic)
                    End If
                Case Else
            End Select
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            Go()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            isFilterByName = True
            pageIndex = 0
            BindPage()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Try
            isFilterByName = False
            pageIndex = 0
            BindPage()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub bsClinic_AddingNew(sender As Object, e As System.ComponentModel.AddingNewEventArgs) Handles bsClinic.AddingNew
        Dim dv As DataView = CType(Me.bsClinic.List, DataView)
        Dim row As DataRowView = dv.AddNew()

        row("IsActive") = True
        e.NewObject = row
    End Sub

    Private Sub btnCancel_MouseEnter(sender As Object, e As EventArgs) Handles btnCancel.MouseEnter
        isEnableValidation = False
    End Sub

    Private Sub btnCancel_MouseLeave(sender As Object, e As EventArgs) Handles btnCancel.MouseLeave
        isEnableValidation = True
    End Sub

    Private Sub btnClose_MouseEnter(sender As Object, e As EventArgs) Handles btnClose.MouseEnter
        isEnableValidation = False
    End Sub

    Private Sub btnClose_MouseLeave(sender As Object, e As EventArgs) Handles btnClose.MouseLeave
        isEnableValidation = True
    End Sub

#Region "Subroutines"

    Private Sub BindPage()
        Try
            totalCount = 0

            If isFilterByName = True Then
                If String.IsNullOrEmpty(txtName.Text.Trim) Then
                    Me.adpClinic.FillDoctorByEmployeeName(Me.dsLeaveFiling.Clinic, pageIndex, pageSize, totalCount, Nothing)
                Else
                    Me.adpClinic.FillDoctorByEmployeeName(Me.dsLeaveFiling.Clinic, pageIndex, pageSize, totalCount, txtName.Text.Trim)
                End If
            Else
                Me.adpClinic.FillDoctor(Me.dsLeaveFiling.Clinic, pageIndex, pageSize, totalCount)
            End If

            Me.bsClinic.DataSource = Me.dsLeaveFiling
            Me.bsClinic.DataMember = dtClinic.TableName
            Me.bsClinic.ResetBindings(True)
            dgvList.AutoGenerateColumns = False
            dgvList.DataSource = Me.bsClinic

            If totalCount Mod pageSize = 0 Then
                If totalCount = 0 Then
                    pageCount = (totalCount / pageSize) + 1
                Else
                    pageCount = totalCount / pageSize
                End If
            Else
                pageCount = Math.Truncate(totalCount / pageSize) + 1
            End If

            'current page index and total number of pages
            txtPageNumber.Text = pageIndex + 1
            txtTotalPageNumber.Text = "of " & CInt(pageCount) & " Page(s)"

            'enables pager
            txtPageNumber.Enabled = True
            txtTotalPageNumber.Enabled = True
            BindingNavigatorMoveFirstItem.Enabled = True
            BindingNavigatorMovePreviousItem.Enabled = True
            BindingNavigatorMoveNextItem.Enabled = True
            BindingNavigatorMoveLastItem.Enabled = True
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Go()
        Try
            If String.IsNullOrEmpty(txtPageNumber.Text) Then
                MessageBox.Show("Page not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPageNumber.Focus()
                Return
            End If

            If CInt(txtPageNumber.Text) > pageCount Then
                MessageBox.Show("Page not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPageNumber.Focus()
                Return
            End If

            If CInt(txtPageNumber.Text) = 0 Then
                MessageBox.Show("Page not found.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPageNumber.Focus()
                Return
            End If

            pageIndex = CInt(txtPageNumber.Text) - 1
            BindPage()
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub RefreshList()
        If dgvList IsNot Nothing AndAlso dgvList.CurrentRow IsNot Nothing Then Me.Invoke(New Action(AddressOf GetScrollingIndex))
        pageIndex = 0
        BindPage()
        If dgvList IsNot Nothing AndAlso dgvList.CurrentRow IsNot Nothing Then Me.Invoke(New Action(AddressOf SetScrollingIndex))
    End Sub

    Private Sub GetScrollingIndex()
        indexScroll = dgvList.FirstDisplayedCell.RowIndex
        indexPosition = dgvList.CurrentRow.Index
    End Sub

    Private Sub SetScrollingIndex()
        dgvList.FirstDisplayedScrollingRowIndex = indexScroll
        If dgvList.Rows.Count > indexPosition Then
            dgvList.Rows(indexPosition).Selected = True
        Else
            dgvList.Rows(indexPosition - 1).Selected = True
        End If
        Me.bsClinic.Position = dgvList.SelectedCells(0).RowIndex
    End Sub

#End Region

End Class