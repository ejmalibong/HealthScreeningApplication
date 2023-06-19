Imports System.Data.SqlClient
Imports Microsoft.Synchronization
Imports Microsoft.Synchronization.Data
Imports Microsoft.Synchronization.Data.SqlServer

Module DataSynchronizer
    Private Sub Initialize(ByVal table As String, ByVal serverConnectionString As String, ByVal clientConnectionString As String)
        Using serverConnection As SqlConnection = New SqlConnection(serverConnectionString)

            Using clientConnection As SqlConnection = New SqlConnection(clientConnectionString)
                Dim scopeDescription As DbSyncScopeDescription = New DbSyncScopeDescription(table)
                Dim tableDescription As DbSyncTableDescription = SqlSyncDescriptionBuilder.GetDescriptionForTable(table, serverConnection)
                scopeDescription.Tables.Add(tableDescription)
                Dim serverProvision As SqlSyncScopeProvisioning = New SqlSyncScopeProvisioning(serverConnection, scopeDescription)
                serverProvision.Apply()
                Dim clientProvision As SqlSyncScopeProvisioning = New SqlSyncScopeProvisioning(clientConnection, scopeDescription)
                clientProvision.Apply()
            End Using
        End Using
    End Sub

    Sub Synchronize(ByVal tableName As String, ByVal serverConnectionString As String, ByVal clientConnectionString As String)
        Initialize(tableName, serverConnectionString, clientConnectionString)
        Synchronize(tableName, serverConnectionString, clientConnectionString, SyncDirectionOrder.DownloadAndUpload)
        CleanUp(tableName, serverConnectionString, clientConnectionString)
    End Sub

    Private Sub Synchronize(ByVal scopeName As String, ByVal serverConnectionString As String, ByVal clientConnectionString As String, ByVal syncDirectionOrder As SyncDirectionOrder)
        Using serverConnection As SqlConnection = New SqlConnection(serverConnectionString)

            Using clientConnection As SqlConnection = New SqlConnection(clientConnectionString)
                Dim agent = New SyncOrchestrator With {
                    .LocalProvider = New SqlSyncProvider(scopeName, clientConnection),
                    .RemoteProvider = New SqlSyncProvider(scopeName, serverConnection),
                    .Direction = syncDirectionOrder
                }
                AddHandler TryCast(agent.RemoteProvider, RelationalSyncProvider).SyncProgress, AddressOf dbProvider_SyncProgress
                AddHandler TryCast(agent.LocalProvider, RelationalSyncProvider).ApplyChangeFailed, AddressOf dbProvider_SyncProcessFailed
                AddHandler TryCast(agent.RemoteProvider, RelationalSyncProvider).ApplyChangeFailed, AddressOf dbProvider_SyncProcessFailed
                agent.Synchronize()
            End Using
        End Using
    End Sub

    Private Sub dbProvider_SyncProcessFailed(ByVal sender As Object, ByVal e As DbApplyChangeFailedEventArgs)

    End Sub

    Private Sub dbProvider_SyncProgress(ByVal sender As Object, ByVal e As DbSyncProgressEventArgs)

    End Sub

    Public Enum DbConflictType
        ErrorsOccured = 0
        LocalUpdateRemoteUpdate = 1
        LocalUpdateRemoteDelete = 2
        LocalDeleteRemoteUpdate = 3
        LocalInsertRemoteInsert = 4
        LocalDeleteRemoteDelete = 5
    End Enum

    Private Sub CleanUp(ByVal scopeName As String, ByVal serverConnectionString As String, ByVal clientConnectionString As String)
        Using serverConnection As SqlConnection = New SqlConnection(serverConnectionString)

            Using clientConnection As SqlConnection = New SqlConnection(clientConnectionString)
                Dim serverDeprovisioning As SqlSyncScopeDeprovisioning = New SqlSyncScopeDeprovisioning(serverConnection)
                Dim clientDeprovisioning As SqlSyncScopeDeprovisioning = New SqlSyncScopeDeprovisioning(clientConnection)
                serverDeprovisioning.DeprovisionScope(scopeName)
                serverDeprovisioning.DeprovisionStore()
                clientDeprovisioning.DeprovisionScope(scopeName)
                clientDeprovisioning.DeprovisionStore()
            End Using
        End Using
    End Sub
End Module
