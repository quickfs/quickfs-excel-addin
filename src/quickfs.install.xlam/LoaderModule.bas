Attribute VB_Name = "LoaderModule"
Option Explicit
Option Private Module

Public loadingManager As Boolean
Public updatingFunctions As Boolean

Public Function IsLoadingManager() As Boolean
    IsLoadingManager = loadingManager
End Function

Public Function IsUpdatingFunctions() As Boolean
    IsUpdatingFunctions = updatingFunctions
End Function

' Check if the functions add-in is installed alongside
Public Function HasAddInFunctions() As Boolean
    HasAddInFunctions = _
        SafeDir(LocalPath(AddInFunctionsFile)) <> "" Or _
        SafeDir(LocalPath(AddInFunctionsFile), vbHidden) <> "" Or _
        HasLegacyFunctions
End Function

' Check if the functions add-in is installed alongside
Public Function HasLegacyFunctions() As Boolean
    HasLegacyFunctions = _
        SafeDir(LocalPath(LegacyFunctionsFile)) <> "" Or _
        SafeDir(LocalPath(LegacyFunctionsFile), vbHidden) <> ""
End Function

' Load the functions add-in installed alongside
Public Sub LoadAddInFunctions()

    ' Make sure add-in is installed on Mac 2011 because
    ' otherwise we'll get a macro prompt every time we
    ' open excel
    If ExcelVersion = "Mac2011" Then
        Dim addin As addin, installed As addin
        For Each addin In Application.AddIns
            If addin.name = AddInFunctionsFile Then
                Set installed = addin
                Exit For
            End If
        Next addin
        
        If addin Is Nothing Then
            Set installed = Application.AddIns.Add(LocalPath(AddInFunctionsFile), True)
        End If
        
        installed.installed = True
    End If

    ' If the functions add-in is already loaded,
    ' we should just exit.
    If LoadedAddInFunctions Then Exit Sub
    
    Dim appSec As MsoAutomationSecurity
    appSec = Application.AutomationSecurity
    
    If HasLegacyFunctions Then
        Name LocalPath(LegacyFunctionsFile) As StagingPath(AddInFunctionsFile)
    End If
    
    ' If an update is staged, promote it to the active
    ' add-in. Only do this if this is an installed add-in
    ' so that we don't accidentally overwrite a dev
    ' version of the functions add-in.
    If HasStagedUpdate And AddInInstalled Then
        PromoteStagedUpdate
    End If
    
    On Error GoTo RemoveAddInFunctions
    
    ' Load the functions add-in
    LogMessage "Loading add-in functions from " & LocalPath(AddInFunctionsFile)
    Application.AutomationSecurity = msoAutomationSecurityLow
    Call Workbooks.Open(LocalPath(AddInFunctionsFile))
    Application.AutomationSecurity = appSec
    LogMessage "Loaded add-in functions v" & AddInVersion(AddInFunctionsFile)
    
    Exit Sub

RemoveAddInFunctions:
    ' If for some reason we can't open the functions
    ' component, the workbook may be corrupted.
    ' Just remove all traces so it will be re-downloaded
    ' on the next restart.
    
    LogMessage "Unable to load add-in functions: " & Err.Description
    
    Application.AutomationSecurity = appSec
    
    RemoveAddInFunctions
    
    MsgBox _
        Title:="[QuickFS] Add-in Error", _
        Prompt:="The QuickFS add-in functions were not loaded correctly. " & _
                "Please try restarting Excel and contact support@quickfs.net if this problem persists.", _
        Buttons:=vbCritical
End Sub

' Ensures that functions add-in is uninstalled and unloaded
Public Function UninstallAddInFunctions() As Boolean
    Dim addin As addin
    For Each addin In Application.AddIns
        If (addin.name = AddInFunctionsFile Or addin.name = LegacyFunctionsFile) And addin.installed Then
            LogMessage "Uninstalling add-in functions"
            addin.installed = False
            UninstallAddInFunctions = True
            Exit Function
        End If
    Next addin
End Function

' Checks if the functions add-in is currently loaded
Public Function LoadedAddInFunctions() As Boolean
    ' The add-in may be loaded as a hidden file, so
    ' it won't always show up in the add-ins list.
    ' So the safest thing to do is check if the workbook
    ' itself is open. If the call below succeeds, then
    ' we know it's loaded.
    On Error Resume Next
    Dim loaded As String
    loaded = ""
    loaded = Workbooks(AddInFunctionsFile).name
    loaded = Workbooks(LegacyFunctionsFile).name
    If loaded <> "" Then
        LoadedAddInFunctions = True
    Else
        LoadedAddInFunctions = False
    End If
End Function

' Unloads the currently loaded functions add-in.
' Does nothing if the add-in is not loaded.
Public Function UnloadAddInFunctions() As Boolean
    UnloadAddInFunctions = UnloadFunctions(AddInFunctionsFile) And UnloadFunctions(LegacyFunctionsFile)
End Function

Private Function UnloadFunctions(name As String) As Boolean
    Dim openName As String, canUnloadFunctions As Boolean
    
    On Error GoTo Unloaded
    
    openName = Workbooks(name).name
    canUnloadFunctions = _
        Not Application.Run(name & "!IsUpdatingManager") And _
        Not Application.Run(name & "!IsCheckingUpdates")
    If Not canUnloadFunctions Then Exit Function

    ' Try to close workbook. If either of these
    ' calls fail it likely means the workbook is
    ' closed.
    LogMessage "Unloading add-in functions v" & AddInVersion(name)
    Workbooks(name).Close
    openName = Workbooks(name).name
    Exit Function
    
Unloaded:
    UnloadFunctions = True
End Function

' Check if staged functions add-in is available
Private Function HasStagedUpdate() As Boolean
    HasStagedUpdate = _
        SafeDir(StagingPath(AddInFunctionsFile)) <> "" Or _
        SafeDir(StagingPath(AddInFunctionsFile), vbHidden) <> "" Or _
        HasStagedLegacyUpdate
End Function

Private Function HasStagedLegacyUpdate() As Boolean
    HasStagedLegacyUpdate = _
        SafeDir(StagingPath(LegacyFunctionsFile)) <> "" Or _
        SafeDir(StagingPath(LegacyFunctionsFile), vbHidden) <> ""
End Function

' Promotes the staged functions add-in to active
Public Sub PromoteStagedUpdate()
    If updatingFunctions Then Exit Sub

    If Not HasStagedUpdate And Not HasLegacyFunctions Then Exit Sub
    
    On Error GoTo Finish
    updatingFunctions = True
    If UnloadAddInFunctions Then
        LogMessage "Promoting staged add-in functions"
        If HasAddInFunctions And Not HasLegacyFunctions Then
            SetAttr LocalPath(AddInFunctionsFile), vbNormal
            Kill LocalPath(AddInFunctionsFile)
        End If
        
        If HasLegacyFunctions And Not HasAddInFunctions Then
            SetAttr LocalPath(LegacyFunctionsFile), vbNormal
            Kill LocalPath(LegacyFunctionsFile)
        ElseIf HasLegacyFunctions Then
            Name LocalPath(LegacyFunctionsFile) As StagingPath(AddInFunctionsFile)
        End If
        
        If HasStagedLegacyUpdate Then
            Name StagingPath(LegacyFunctionsFile) As StagingPath(AddInFunctionsFile)
        End If
        
        Name StagingPath(AddInFunctionsFile) As LocalPath(AddInFunctionsFile)
        VBA.SetAttr LocalPath(AddInFunctionsFile), vbHidden
        
        #If Mac Then
            MsgBox _
                Title:="[QuickFS] Add-In Functions Updated", _
                Prompt:="A new version of the add-in functions have been installed. " & _
                        "You may be prompted to enable the updated macros. " & _
                        "Macros must be enabled or the add-in will not function properly."
        #End If
        
        LoadAddInFunctions
    End If
    
Finish:
    updatingFunctions = False
End Sub

