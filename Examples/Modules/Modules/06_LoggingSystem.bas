==========================================================
' VBA Enterprise Framework - Logging System Module
'==========================================================
Option Explicit

'Logging Configuration
Private Type LogConfig
    LogLevel As String
    LogPath As String
    MaxLogSize As Long
    RotateLog As Boolean
End Type

'Logging Manager
Public Sub LogEvent(eventType As String, _
                   message As String, _
                   Optional details As Variant)
                   
    'Create logging manager
    Dim logMgr As New LogManager
    
    With logMgr
        'Process log entry
        .PrepareLogEntry
        .AddEventDetails eventType, message
        .AppendAdditionalInfo details
        .WriteLog
        
        'Check log size
        If mLogConfig.RotateLog Then
            .CheckLogSize
            .RotateIfNeeded
        End If
    End With
End Sub

'Log File Manager
Private Sub ManageLogFile()
    'Initialize file manager
    Dim fileMgr As New LogFileManager
    
    With fileMgr
        'Manage log file
        .CheckFileSize
        .ArchiveIfNeeded
        .CleanupOldLogs
        .OptimizeStorage
    End With
End Sub

'Log Analysis
Public Function AnalyzeLogs(dateRange As String) As Dictionary
    'Initialize analyzer
    Dim analyzer As New LogAnalyzer
    
    With analyzer
        'Analyze logs
        .LoadLogFiles dateRange
        .ProcessEntries
        .GenerateStatistics
        .PrepareReport
        
        Set AnalyzeLogs = .GetResults
    End With
End Function
