Imports BlackCoffeeLibrary
Imports System.Data.SqlClient
Imports System.Deployment.Application

Public Class frmLogin
    Private connection As New clsConnection
    Private dbScreening As New SqlDbMethod(connection.ServerConnection)
    Private dbJeonsoft As New SqlDbMethod(connection.JeonsoftConnection)
    Private dbMain As New Main

    Private employeeId As Integer = 0
    Private employeeCode As String = String.Empty
    Private employeeName As String = String.Empty
    Private positionName As String = String.Empty

    Private isAdmin As Boolean = False

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ApplicationDeployment.IsNetworkDeployed Then
            lblVersion.Text = "ver. " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Else
            lblVersion.Text = "ver. " & Application.ProductVersion.ToString
        End If

        Me.ActiveControl = txtEmployeeId
    End Sub

    Private Sub frmLogin_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        Me.ActiveControl = txtEmployeeId
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Try
            If ApplicationDeployment.IsNetworkDeployed Then
                If Not My.Computer.Network.IsAvailable Then
                    MessageBox.Show("No network connection.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            If String.IsNullOrEmpty(txtEmployeeId.Text.Trim) Then
                MessageBox.Show("Please enter your employee ID.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtEmployeeId.Focus()
                Return
            End If

            If String.IsNullOrEmpty(txtPassword.Text.Trim) Then
                MessageBox.Show("Please enter your password.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPassword.Focus()
                Return
            End If

            Dim count1 As Integer = 0
            Dim prmLogin1(1) As SqlParameter
            prmLogin1(0) = New SqlParameter("@EmployeeCode", SqlDbType.NVarChar)
            prmLogin1(0).Value = txtEmployeeId.Text.Trim
            prmLogin1(1) = New SqlParameter("@Password", SqlDbType.NVarChar)
            prmLogin1(1).Value = txtPassword.Text.Trim

            'check if nbc nurses
            'use latin1 general collation for case-sensitive password
            count1 = dbScreening.ExecuteScalar("SELECT COUNT(EmployeeId) FROM VwEmployee WHERE EmployeeCode = @EmployeeCode AND " &
                                               "TRIM(Password) COLLATE Latin1_General_CS_AS = @Password AND IsActive = 1 AND EmployeeId IN ( " &
                                               "SELECT EmployeeId FROM Nurse)", CommandType.Text, prmLogin1)

            If count1 > 0 Then 'nbc nurses
                Dim prmLogin2(1) As SqlParameter
                prmLogin2(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                prmLogin2(0).Value = txtEmployeeId.Text.Trim
                prmLogin2(1) = New SqlParameter("@Password", SqlDbType.NVarChar)
                prmLogin2(1).Value = txtPassword.Text.Trim

                Using reader As IDataReader = dbScreening.ExecuteReader("RdEmployee", CommandType.StoredProcedure, prmLogin2)
                    GetUserInformation(reader)
                End Using
            Else 'non-nbc doctors
                Dim count2 As Integer = 0
                Dim prmLogin3(1) As SqlParameter
                prmLogin3(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                prmLogin3(0).Value = txtEmployeeId.Text.Trim
                prmLogin3(1) = New SqlParameter("@Password", SqlDbType.NVarChar)
                prmLogin3(1).Value = txtPassword.Text.Trim

                count2 = dbScreening.ExecuteScalar("SELECT COUNT(EmployeeId) FROM VwClinic WHERE EmployeeCode = @EmployeeCode AND " &
                                                   "TRIM(Password) COLLATE Latin1_General_CS_AS = @Password AND IsActive = 1", CommandType.Text, prmLogin3)

                If count2 > 0 Then 'non-nbc doctors
                    Dim prmLogin4(1) As SqlParameter
                    prmLogin4(0) = New SqlParameter("@EmployeeCode", SqlDbType.VarChar)
                    prmLogin4(0).Value = txtEmployeeId.Text.Trim
                    prmLogin4(1) = New SqlParameter("@Password", SqlDbType.NVarChar)
                    prmLogin4(1).Value = txtPassword.Text.Trim

                    Using reader As IDataReader = dbScreening.ExecuteReader("RdClinic", CommandType.StoredProcedure, prmLogin4)
                        GetUserInformation(reader)
                    End Using
                Else 'unauthorized login - incorrect credentials or inactive user
                    MessageBox.Show("Incorrect employeee ID or password.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtEmployeeId.Focus()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(dbMain.SetExceptionMessage(ex), "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Application.Exit()
    End Sub

#Region "Subroutines"

    Private Sub GetUserInformation(reader As IDataReader)
        While reader.Read
            employeeId = reader.Item("EmployeeId")
            employeeCode = reader.Item("EmployeeCode").ToString.Trim
            employeeName = reader.Item("EmployeeName").ToString.Trim
            positionName = reader.Item("PositionName").ToString.Trim
            isAdmin = reader.Item("IsAdmin")
        End While
        reader.Close()

        Me.Hide()

        Dim frmScreenList As New frmScreenList(employeeId, employeeCode, employeeName, positionName, isAdmin)
        frmScreenList.Show()
        txtEmployeeId.Clear()
        txtPassword.Clear()
    End Sub

#End Region

End Class