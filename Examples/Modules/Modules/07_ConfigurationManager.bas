==========================================================
' VBA Enterprise Framework - Configuration Manager Module
'==========================================================
Option Explicit

'Configuration Settings
Private Type ConfigSettings
    Environment As String
    Settings As Dictionary
    Cached As Boolean
    AutoRefresh As Boolean
End Type

'Configuration Manager
Public Sub LoadConfiguration()
    'Create config manager
    Dim configMgr As New ConfigurationManager
    
    With configMgr
        'Load configuration
        .LoadEnvironmentSettings
        .ValidateSettings
        .ApplyConfiguration
        .CacheIfEnabled
        
        'Setup auto refresh
        If mConfigSettings.AutoRefresh Then
            .SetupRefreshTimer
            .MonitorChanges
        End If
    End With
End Sub

'Settings Handler
Public Function UpdateSetting(key As String, _
                            value As Variant) As Boolean
    'Initialize handler
    Dim handler As New SettingsHandler
    
    With handler
        'Process update
        .ValidateSetting key, value
        .BackupCurrentValue
        .ApplyChange
        .VerifyUpdate
        
        UpdateSetting = .UpdateSuccess
    End With
End Function

'Configuration Validator
Private Function ValidateConfig() As Boolean
    'Initialize validator
    Dim validator As New ConfigValidator
    
    With validator
        'Validate configuration
        .CheckRequiredSettings
        .ValidateDataTypes
        .VerifyDependencies
        .LogValidation
        
        ValidateConfig = .ValidationPassed
    End With
End Function
