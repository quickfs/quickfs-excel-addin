Attribute VB_Name = "InstallerModule"
Option Explicit
Option Private Module

Public installing As Boolean
Public uninstalling As Boolean

Public Function IsInstalling() As Boolean
    IsInstalling = installing Or uninstalling
End Function

' Install this workbook as an add-in in the default
' Excel add-in location. This simplifies deployment
' and management across different platforms and
' ensures valid trust settings for the add-in.
'
' This function returns `True` if it is called from
' an already-installed instance of the add-in,
' otherwise it will return `False`.

Public Function InstallAddIn(self As Workbook) As Boolean
    ' Don't run if add-in is already installed
    InstallAddIn = (VBA.LCase(self.name) = VBA.LCase(AddInInstalledFile))
    #If Mac Then
        If Not InstallAddIn Then MacInstallPrompt.Show
    #Else
        If Not InstallAddIn Then InstallPrompt.Show
    #End If
End Function

Public Sub FinishInstalling()
    On Error GoTo HandleError
    
    installing = True
    Dim installPath As String
    installPath = SavePath(AddInInstalledFile)
    
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If VBA.LCase(i.name) = VBA.LCase(AddInInstalledFile) Then
            i.installed = False
            On Error Resume Next
            Workbooks(i.name).Close
            If SafeDir(i.FullName) <> "" Then Kill i.FullName
            If SafeDir(i.FullName, vbHidden) <> "" Then
                SetAttr i.FullName, vbNormal
                Kill i.FullName
            End If
            If VBA.LCase(i.FullName) = VBA.LCase(installPath) Then Set installed = i
        End If
    Next i
    
    ' Make sure any existing installation is removed
    ' because Mac will not overwrite existing file
    On Error Resume Next
    Workbooks(AddInInstalledFile).Close
    If SafeDir(installPath) <> "" Then Kill installPath
    If SafeDir(installPath, vbHidden) <> "" Then
        SetAttr installPath, vbNormal
        Kill installPath
    End If
    On Error GoTo HandleError
    
    ' Copy the workbook into the default add-in location
    ' and remove any existing functions component. The
    ' corresponding functions component will be installed
    ' automatically
    SaveCopy ThisWorkbook, installPath
    RemoveAddInFunctions
    
    ' If there is a local version of the
    ' quickfs.functions.xlam add-in, we
    ' install that. This is primarily for
    ' convenient installation of dev
    ' (e.g. non-released) add-in versions
    If HasAddInFunctions And IsDevDir Then
        VBA.FileCopy LocalPath(AddInFunctionsFile), SavePath(AddInFunctionsFile)
        VBA.SetAttr SavePath(AddInFunctionsFile), vbHidden
    Else
        InstallAddInFunctions
    End If
    
    ' Add the workbook as an add-in if this is a
    ' new installation or if the path has changed
    If installed Is Nothing Then
        Dim wb As Workbook
        
        ' AddIns.Add will fail unless a workbook is open
        ' so we create a hidden one here and clean up after
        If Application.Workbooks.count = 0 Then
            Application.ScreenUpdating = False
            Set wb = Application.Workbooks.Add
        End If
        
        Set installed = Application.AddIns.Add(installPath, True)
        
        If Not wb Is Nothing Then wb.Close
    End If
    
    ' Activate the installed add-in
    installed.installed = True
    
    ' Our work is done! Close the installer
    ' workbook since the in-place add-in is
    ' now running
    Application.ScreenUpdating = True
    installing = False
    
    ' Warn about restarts
    Dim leftover As addin
    For Each i In Application.AddIns
        If (VBA.LCase(i.name) = VBA.LCase(AddInInstalledFile)) And (VBA.LCase(i.FullName) <> VBA.LCase(installPath)) Then
            Set leftover = i
            Exit For
        End If
    Next i
    
    If leftover Is Nothing Then
        MsgBox _
            Title:="[QuickFS] Installation Succeeded", _
            Prompt:="The QuickFS add-in is now installed and ready to use! Enjoy!", _
            Buttons:=vbInformation
    Else
        MsgBox _
            Title:="[QuickFS] Restart Required", _
            Prompt:="Excel must be restarted to continue the installation. " & _
                    "You may be required to restart Excel once more before the installation is complete.", _
            Buttons:=vbInformation
        Application.Quit
        ThisWorkbook.Close
    End If
    
    On Error Resume Next
    ThisWorkbook.Close
    Exit Sub
    
HandleError:
    installing = False
    Application.ScreenUpdating = True
    LogMessage "Installation Error: " & Err.Description
    MsgBox _
        Title:="[QuickFS] Add-in Error", _
        Prompt:="Unable to install the QuickFS add-on. Please try again and contact support@quickfs.net if this problem persists.", _
        Buttons:=vbCritical
End Sub

Public Sub CancelInstall()
    If IsDevDir Then
        ' If we're running this from a development directory,
        ' close the installed add-ins and continue
        Dim i As addin
        For Each i In Application.AddIns
            If VBA.LCase(i.name) = VBA.LCase(AddInInstalledFile) Then
                ' Originally wanted to use AddIn.IsOpen here, but that
                ' seems to not be available on Mac so we have to just
                ' try to close the workbook directly and ignore errors
                On Error Resume Next
                Workbooks(i.name).Close
                UnloadAddInFunctions
                LoadAddInFunctions
                Exit For
            End If
        Next i
    Else
        ' This add-in shouldn't be run outside
        ' of the installation directory
        LogMessage "Installation canceled"
        ThisWorkbook.Close
    End If
End Sub

Public Sub UninstallAddIn()
    LogMessage "Uninstalling add-in"
    
    uninstalling = True
    
    On Error Resume Next
        
    ' Uninstall and delete installed add-in files
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If VBA.InStr(VBA.LCase(i.name), "quickfs") > 0 Then
            i.installed = False
            Workbooks(i.name).Close
            If SafeDir(i.FullName) <> "" Then Kill i.FullName
            If SafeDir(i.FullName, vbHidden) <> "" Then
                SetAttr i.FullName, vbNormal
                Kill i.FullName
            End If
        End If
    Next i
    
    cd SavePath
    
    ' Second check to make sure the add-in manager is removed
    Workbooks(AddInInstalledFile).Close
    If SafeDir(LocalPath(AddInInstalledFile)) <> "" Then Kill LocalPath(AddInInstalledFile)
    If SafeDir(LocalPath(AddInInstalledFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInInstalledFile), vbNormal
        Kill LocalPath(AddInInstalledFile)
    End If
    
    If SafeDir(StagingPath(AddInInstalledFile)) <> "" Then Kill StagingPath(AddInInstalledFile)
    If SafeDir(StagingPath(AddInInstalledFile), vbHidden) <> "" Then
        SetAttr StagingPath(AddInInstalledFile), vbNormal
        Kill StagingPath(AddInInstalledFile)
    End If

    ' Second check to make sure the add-in functions are removed
    Workbooks(AddInFunctionsFile).Close
    If SafeDir(LocalPath(AddInFunctionsFile)) <> "" Then Kill LocalPath(AddInFunctionsFile)
    If SafeDir(LocalPath(AddInFunctionsFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInFunctionsFile), vbNormal
        Kill LocalPath(AddInFunctionsFile)
    End If
    
    If SafeDir(StagingPath(AddInFunctionsFile)) <> "" Then Kill StagingPath(AddInFunctionsFile)
    If SafeDir(StagingPath(AddInFunctionsFile), vbHidden) <> "" Then
        SetAttr StagingPath(AddInFunctionsFile), vbNormal
        Kill StagingPath(AddInFunctionsFile)
    End If
    
    ' Delete the api key file
    If SafeDir(LocalPath(AddInKeyFile)) <> "" Then Kill LocalPath(AddInKeyFile)
    If SafeDir(LocalPath(AddInKeyFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInKeyFile), vbNormal
        Kill LocalPath(AddInKeyFile)
    End If
    
    ' Delete the config file
    If SafeDir(LocalPath(AddInSettingsFile)) <> "" Then Kill LocalPath(AddInSettingsFile)
    If SafeDir(LocalPath(AddInSettingsFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInSettingsFile), vbNormal
        Kill LocalPath(AddInSettingsFile)
    End If
    
    cd ThisWorkbook.path
    
    uninstalling = False
    
    MsgBox _
        Title:="[QuickFS] Add-In Removed", _
        Prompt:="The QuickFS add-in has been successfully removed. Hope to see you back soon!", _
        Buttons:=vbInformation
    
    LogMessage "Add-in uninstalled"
    
    ThisWorkbook.Close
End Sub

Public Sub InstallAddInFunctions()
    cd SavePath
    
    On Error GoTo HandleError
    DownloadFile DOWNLOADS_URL & "/v" & AddInVersion & "/" & AddInFunctionsFile, StagingPath(AddInFunctionsFile)
    VBA.SetAttr StagingPath(AddInFunctionsFile), vbHidden
    
    LogMessage "Add-in functions v" & AddInVersion & " have been downloaded and staged"
    
    PromoteStagedUpdate
    
    cd ThisWorkbook.path
    
    Exit Sub
HandleError:
    On Error Resume Next
    
    LogMessage "Unable to install add-in functions: " & Err.Description
    
    MsgBox _
        Title:="[QuickFS] Installation Failed", _
        Prompt:="The add-in functions could not be installed at this time. Please try again and contact support@quickfs.net if this problem persists.", _
        Buttons:=vbCritical
        
    RemoveAddInFunctions
    
    cd ThisWorkbook.path
End Sub

Public Sub RemoveAddInFunctions()
    cd SavePath
    
    On Error Resume Next
    UninstallAddInFunctions
    UnloadAddInFunctions
    
    SetAttr LocalPath(AddInFunctionsFile), vbNormal
    Kill LocalPath(AddInFunctionsFile)
    
    SetAttr StagingPath(AddInFunctionsFile), vbNormal
    Kill StagingPath(AddInFunctionsFile)
    
    LogMessage "Removed add-in functions workbook"
    
    cd ThisWorkbook.path
End Sub

Public Sub CloseInstaller()
    If VBA.LCase(ThisWorkbook.name) = VBA.LCase(AddInInstallerFile) Then Exit Sub
    On Error GoTo Closed
    Dim opened As String
    opened = Workbooks(AddInInstallerFile).name
    If Not Application.Run(AddInInstallerFile & "!IsInstalling") Then
        Workbooks(AddInInstallerFile).Close
    End If
Closed:
End Sub

Function SavePath(Optional file As String)
   #If Mac Then
        If ExcelVersion = "Mac2016" Then
            SavePath = MacScript("return POSIX path of (path to home folder) as string")
            Dim Path15 As String, Path16 As String
            Path15 = SavePath & "Library/Application Support/Microsoft/AppData/Office/15.0"
            Path16 = SavePath & "Library/Application Support/Microsoft/Office/16.0"
            If SafeDir(Path16, vbDirectory) <> "" Then
                SavePath = Path16 & "/Add-Ins/"
            ElseIf SafeDir(Path15, vbDirectory) <> "" Then
                SavePath = Path15 & "/Add-Ins/"
            Else
                SavePath = Application.UserLibraryPath
            End If
        Else
            SavePath = Application.LibraryPath
        End If
    #Else
        SavePath = Application.UserLibraryPath
    #End If
    If file <> "" Then SavePath = SavePath & file
End Function

Sub SaveCopy(wb, path As String)
    If SafeDir(path) <> "" Then Kill path
    If SafeDir(path, vbHidden) <> "" Then
        SetAttr path, vbNormal
        Kill path
    End If
    
    wb.SaveCopyAs path
End Sub

Function IsDevDir() As Boolean
    IsDevDir = SafeDir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> ""
End Function

' Changing the location of an existing add-in causes serious
' problems in Excel. Since the Mac2016 version was originally
' installed in a different location, we need to do some cleanup
' of database entries that may get out of sync.

Sub CleanUpUninstalledAddIns()
    Dim installPath As String
    installPath = SavePath(AddInInstalledFile)
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If (VBA.LCase(i.name) = VBA.LCase(AddInInstalledFile)) And Not i.installed And VBA.LCase(i.FullName) <> VBA.LCase(installPath) Then
            ClearAddInRegKey i.FullName
        End If
    Next i
End Sub

Sub ClearAddInRegKey(path As String)
    If ExcelVersion = "Mac2016" Then
        Dim cmd As String, result As ShellResult
        cmd = "echo 'DELETE FROM HKEY_CURRENT_USER_values WHERE value='""'""'""" & path & """'""'""';' | sqlite3 '" & Mac2016Registry & "'"
        Debug.Print cmd
        result = xHelpersWeb.ExecuteInShell(cmd)
        cmd = "echo 'DELETE FROM HKEY_CURRENT_USER_values WHERE name=""" & path & """;' | sqlite3 '" & Mac2016Registry & "'"
        Debug.Print cmd
        result = xHelpersWeb.ExecuteInShell(cmd)
        ' MsgBox _
        '     Title:="[QuickFS] Installation Succeeded", _
        '     Prompt:="The installation succeeded, but Excel must be restarted twice to remove all traces of the previous version. " & _
        '             "This should not be necessary for future updates. Click OK to exit Excel now.", _
        '     Buttons:=vbInformation
        ' Call Application.Workbooks.Add
        ' Application.Quit
        ' ThisWorkbook.Close
        ' Exit Sub
    End If
End Sub

Function Mac2016Registry() As String
    If ExcelVersion = "Mac2016" Then
        Mac2016Registry = MacScript("return POSIX path of (path to desktop folder) as string")
        Mac2016Registry = Replace(Mac2016Registry, "/Desktop", "") & "Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB.reg"
    End If
End Function
