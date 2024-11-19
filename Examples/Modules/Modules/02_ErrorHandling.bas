
'==========================================================
' VBA Enterprise Framework - Error Handling Module
'==========================================================
Option Explicit

'Error Handler Configuration
Private Type ErrorConfig
    LogErrors As Boolean
    ErrorLog As String
    NotifyAdmin As Boolean
    DetailLevel As String
End Type

'Central Error Handler
Public Sub HandleError(ByVal errorNumber As Long, _
                      ByVal errorDescription As String, _
                      ByVal moduleName As String, _
                      ByVal procedureName As String)
                      
    'Create error manager
    Dim errorMgr As New ErrorManager
    
    With errorMgr
        'Process error
        .LogError errorNumber, errorDescription
        .NotifySupport moduleName, procedureName
        .AttemptRecovery
        .UpdateErrorLog
    End With
End Sub

'Error Logging Function
Private Function LogError(errorDetails As Variant) As Boolean
    'Initialize logger
    Dim logger As New ErrorLogger
    
    With logger
        'Log error details
        .PrepareLog
        .AddErrorDetails errorDetails
        .SaveLog
        .NotifyIfCritical
        
        LogError = .LogSuccess
    End With
End Function
