

'==========================================================
' VBA Enterprise Framework - Foundation Module
'==========================================================
Option Explicit

'Example: Basic Configuration Structure
Private Type FrameworkConfig
    Modules As Collection
    Settings As Dictionary
    Logging As Boolean
    Debug As Boolean
End Type

'Example: Framework Initialization
Public Sub InitializeFramework()
    'Create framework manager
    Dim fwMgr As New FrameworkManager
    
    With fwMgr
        'Configure framework
        .LoadModules
        .InitializeSettings
        .EnableLogging
        .StartFramework
    End With
End Sub
