

'==========================================================
' VBA Enterprise Framework - Database Connection Module
'==========================================================
Option Explicit

'Database Configuration
Private Type DBConfig
    ConnectionString As String
    Timeout As Long
    RetryAttempts As Integer
    UseTransaction As Boolean
End Type

'Database Connection Manager
Public Function CreateConnection() As Boolean
    'Create connection manager
    Dim connMgr As New ConnectionManager
    
    With connMgr
        'Setup connection
        .LoadConfiguration
        .InitializeConnection
        .TestConnection
        .EnableTransaction
        
        CreateConnection = .Connected
    End With
End Function

'Query Execution Handler
Public Function ExecuteQuery(sqlQuery As String) As Recordset
    'Initialize query handler
    Dim queryHandler As New QueryHandler
    
    With queryHandler
        'Process query
        .ValidateQuery sqlQuery
        .PrepareExecution
        .ExecuteSQL
        .HandleResults
        
        Set ExecuteQuery = .ResultSet
    End With
End Function
