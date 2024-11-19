

'==========================================================
' VBA Enterprise Framework - System Monitor Module
'==========================================================
Option Explicit

'Monitor Configuration
Private Type MonitorConfig
    Metrics As Collection
    Alerts As Dictionary
    Logging As Dictionary
    RealTime As Boolean
End Type

'System Monitor Manager
Public Sub InitializeSystemMonitor()
    'Create monitor manager
    Dim monitorMgr As New SystemMonitorManager
    
    With monitorMgr
        'Setup monitoring
        .LoadMetrics mMonitorConfig.Metrics
        .ConfigureAlerts mMonitorConfig.Alerts
        .SetupLogging mMonitorConfig.Logging
        
        'Start monitoring
        .InitializeMonitors
        .StartTracking
        .ProcessAlerts
        .GenerateReports
    End With
End Sub
