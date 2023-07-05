Private handler As AppEventHandler

Private Sub Workbook_Open()
    SafeMkDir SavePath

    LogMessage "Loading add-in manager"
    
    AddInUninstalling = False
    loadingManager = True
    
    ' Install AddIn on first open
    AddInInstalled = InstallAddIn(Me)
    
    ' Download functions component if not available
    If AddInInstalled And Not HasAddInFunctions Then
        InstallAddInFunctions
    End If
    
    ' Download newer version if available and auto-update is not disabled
    If AddInInstalled And GetSetting("autoUpdate", False) And GetSetting("updateOnLaunch", False) Then
        Call DownloadUpdates(blockEvents:=True)
    End If
    
    ' Load the QFS functions
    Call LoadAddInFunctions
    
    Set handler = New AppEventHandler
    
    loadingManager = False
    
    TrimLog
    
    LogMessage "Add-in manager is ready"
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error GoTo Finish
    If SafeDir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> "" Then
'        Dim openName As String
'        openName = Workbooks("vbaDeveloper.xlam").name
'        Call Application.Run("vbaDeveloper.xlam!exportVbProject", "quickfsManager")
    End If
Finish:
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' The functions add-in should NOT be formally installed
    ' (as indicated by a checkmark in the add-ins window),
    ' since it will be loaded by the manager add-in.
    ' If it is installed, we must uninstall it to
    ' avoid problems on startup.
    ' Theoretically it should be very unlikely that the
    ' functions add-in actually ends up in an installed
    ' state since it is a hidden file and won't show up
    ' in the add-ins list, but it doesn't hurt to check.
    '
    ' Note: Mac2011 is the only exception to this. We
    ' actually DO need to install it on Mac2011 or we
    ' get macro prompts every time excel is opened.
    
    On Error Resume Next
    
    If ExcelVersion <> "Mac2011" Then
        Call UninstallAddInFunctions
        Unload InstallPrompt
        Unload MacInstallPrompt
    End If
    
    TrimLog
    
    LogMessage "Unloaded add-in manager"
End Sub