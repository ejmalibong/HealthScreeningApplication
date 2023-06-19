
Namespace My
    
    'This class allows you to handle specific events on the settings class:
    ' The SettingChanging event is raised before a setting's value is changed.
    ' The PropertyChanged event is raised after a setting's value is changed.
    ' The SettingsLoaded event is raised after the setting values are loaded.
    ' The SettingsSaving event is raised before the setting values are saved.
    Partial Friend NotInheritable Class MySettings

        Private Sub MySettings_SettingsLoaded(sender As Object, e As System.Configuration.SettingsLoadedEventArgs) Handles Me.SettingsLoaded
            If SickLeaveScreening.My.MySettings.Default.IsDebug = True Then
                If Environment.MachineName.ToString = "NBCP-DT-032" Then
                    Me.Item("LeaveConnectionString") = "Data Source=NBCP-DT-032\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                    Me.Item("LeaveConnectionStringRpt") = "Data Source=NBCP-DT-032\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                    Me.Item("JeonsoftConnectionString") = "Data Source=NBCP-DT-032\SQLEXPRESS;Initial Catalog=NBCTECHDB;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                Else
                    Me.Item("LeaveConnectionString") = "Data Source=NBCP-LT-043\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                    Me.Item("LeaveConnectionStringRpt") = "Data Source=NBCP-LT-043\SQLEXPRESS;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                    Me.Item("JeonsoftConnectionString") = "Data Source=NBCP-LT-043\SQLEXPRESS;Initial Catalog=NBCTECHDB;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                End If
            Else
                Me.Item("LeaveConnectionString") = "Data Source=LENOVO-AX3RONG2;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                Me.Item("LeaveConnectionStringRpt") = "Data Source=LENOVO-AX3RONG2;Initial Catalog=LeaveFiling;Persist Security Info=True;User ID=sa;Password=Nbc12#"
                Me.Item("JeonsoftConnectionString") = "Data Source=LENOVO-AX3RONG2;Initial Catalog=NBCTECHDB;Persist Security Info=True;User ID=sa;Password=Nbc12#"
            End If
        End Sub

    End Class

End Namespace
