Public Class clsConnection

    Public Function ClientConnection() As String
        If SickLeaveScreening.My.MySettings.Default.IsDebug = True Then
            If Environment.MachineName.ToString.Trim = "NBCP-DT-032" Then
                Return "Data Source=NBCP-DT-032\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            Else
                Return "Data Source=NBCP-LT-043\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            End If
        Else
            Return "Data Source=LENOVO-AX3RONG2;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
        End If
    End Function

    Public Function ServerConnection() As String
        If SickLeaveScreening.My.MySettings.Default.IsDebug = True Then
            If Environment.MachineName.ToString.Trim = "NBCP-DT-032" Then
                Return "Data Source=NBCP-DT-032\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            Else
                Return "Data Source=NBCP-LT-043\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            End If
        Else
            Return "Data Source=LENOVO-AX3RONG2;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
        End If
    End Function

    Public Function JeonsoftConnection() As String
        If SickLeaveScreening.My.MySettings.Default.IsDebug = True Then
            If Environment.MachineName.ToString.Trim = "NBCP-DT-032" Then
                Return "Data Source=NBCP-DT-032\SQLEXPRESS;Initial Catalog=NBCTECHDB;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            Else
                Return "Data Source=NBCP-LT-043\SQLEXPRESS;Initial Catalog=NBCTECHDB;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            End If
        Else
            Return "Data Source=LENOVO-AX3RONG2;Initial Catalog=NBCTECHDB;Persist Security Info=True;User ID=sa;Password=Nbc12#"
        End If
    End Function

End Class