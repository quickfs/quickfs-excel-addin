Attribute VB_Name = "ConfigModule"
Option Explicit
Option Private Module

Public Const RELEASES_URL = "https://api.github.com/repos/quickfs/quickfs-excel-addin/releases"
Public Const DOWNLOADS_URL = "https://github.com/quickfs/quickfs-excel-addin/releases/download"

Public Const AddInInstalledFile = "quickfs.xlam"
Public Const LegacyInstalledFile = "quickfsnet.xlam"
Public Const AddInInstallerFile = "quickfs.install.xlam"
Public Const AddInFunctionsFile = "quickfs.functions.xlam"
Public Const LegacyFunctionsFile = "quickfsnet.functions.xlam"
Public Const AddInKeyFile = "quickfs.key"
Public Const AddInSettingsFile = "quickfs.cfg"
Public Const AddInLogFile = "quickfs.log"

' These will be loaded on Workbook_Open
Public AddInInstalled As Boolean
Public cwd As String
Public AddInUninstalling As Boolean

Public Function AddInManagerFile() As String
    AddInManagerFile = ThisWorkbook.name
End Function

Public Function StagingFile(file As String) As String
    StagingFile = VBA.Left(file, VBA.InStrRev(file, ".")) & "staged" & VBA.Mid(file, InStrRev(file, "."))
End Function

Public Function StagingPath(file As String) As String
    StagingPath = LocalPath(StagingFile(file))
End Function

Public Sub cd(path As String)
    If VBA.Right(path, 1) = Application.PathSeparator Then
        cwd = VBA.Left(path, VBA.Len(path) - 1)
    Else
        cwd = path
    End If
End Sub

Public Function LocalPath(file As String) As String
    If cwd = "" Then cwd = ThisWorkbook.path
    LocalPath = cwd & Application.PathSeparator & file
End Function

Public Function AddInVersion(Optional file As String) As String
    If file = "" Then file = ThisWorkbook.name
    On Error Resume Next
    AddInVersion = Workbooks(file).Sheets("quickfs").Range("AppVersion").value
    AddInVersion = Workbooks(file).Sheets("quickfsnet").Range("AppVersion").value
End Function

Public Function AddInReleaseDate(Optional file As String) As Date
    If file = "" Then file = ThisWorkbook.name
    On Error Resume Next
    AddInReleaseDate = VBA.Now()
    AddInReleaseDate = Workbooks(file).Sheets("quickfs").Range("ReleaseDate").value
    AddInReleaseDate = Workbooks(file).Sheets("quickfsnet").Range("ReleaseDate").value
End Function

Public Function AddInLocation(Optional file As String) As String
    If file = "" Then file = ThisWorkbook.name
    On Error Resume Next
    AddInLocation = Workbooks(file).FullName
End Function

Public Function SafeDir(file As String, Optional attributes As VbFileAttribute) As String
    On Error Resume Next
    SafeDir = VBA.Dir(file, attributes)
End Function

Public Sub SafeMkDir(path As String)
    Dim folder As String
    folder = path
    If Right(path, 1) = Application.PathSeparator Then
        folder = Left(path, Len(path) - 1)
    End If
    If SafeDir(folder, vbDirectory) = vbNullString Then
        #If Mac Then
            Dim appleScript As String
            appleScript = "do shell script ""mkdir -p '" & folder & "'"""
            MacScript appleScript
        #Else
            VBA.MkDir folder
        #End If
    End If
End Sub

Sub auto_add()
End Sub
Sub auto_remove()
    AddInUninstalling = True
End Sub
