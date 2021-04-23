Attribute VB_Name = "UtilitiesModule"
Option Explicit
Option Private Module

Function VersionAtLeast(version As String) As Boolean
    Dim major As Integer, minor As Integer, patch As Integer
    
    ' Get major version
    major = Val(Split(version, ".")(0))
    
    ' Get minor version
    If CountCharacters(version, ".") > 0 Then
        minor = Val(Split(version, ".")(1))
    Else
        minor = 0
    End If
    
    ' Get patch version
    If CountCharacters(version, ".") > 1 Then
        patch = Val(Split(version, ".")(2))
    Else
        patch = 0
    End If
    
    If MajorVersion > major Then
        VersionAtLeast = True
    ElseIf MajorVersion = major And MinorVersion > minor Then
        VersionAtLeast = True
    ElseIf MajorVersion = major And MinorVersion = minor And PatchVersion >= patch Then
        VersionAtLeast = True
    Else
        VersionAtLeast = False
    End If
End Function

Function VersionLessThan(version As String) As Boolean
    Dim major As Integer, minor As Integer, patch As Integer
    
    ' Get major version
    major = Val(Split(version, ".")(0))
    
    ' Get minor version
    If CountCharacters(version, ".") > 0 Then
        minor = Val(Split(version, ".")(1))
    Else
        minor = 0
    End If
    
    ' Get patch version
    If CountCharacters(version, ".") > 1 Then
        patch = Val(Split(version, ".")(2))
    Else
        patch = 0
    End If
    
    If MajorVersion < major Then
        VersionLessThan = True
    ElseIf MajorVersion = major And MinorVersion < minor Then
        VersionLessThan = True
    ElseIf MajorVersion = major And MinorVersion = minor And PatchVersion < patch Then
        VersionLessThan = True
    Else
        VersionLessThan = False
    End If
End Function

Function MajorVersion() As Double
    MajorVersion = Val(Split(Application.version, ".")(0))
End Function

Function MinorVersion() As Double
    If CountCharacters(Application.version, ".") > 0 Then
        MinorVersion = Val(Split(Application.version, ".")(1))
    Else
        MinorVersion = 0
    End If
End Function

Function PatchVersion() As Double
    If CountCharacters(Application.version, ".") > 1 Then
        PatchVersion = Val(Split(Application.version, ".")(2))
    Else
        PatchVersion = 0
    End If
End Function

Function CountCharacters(str As String, char As String) As Long
    Dim i As Integer, count As Integer
    count = 0
    For i = 1 To Len(str)
        If Mid(str, i, 1) = char Then count = count + 1
    Next
    CountCharacters = count
End Function

Public Function ExcelVersion() As String
    Dim version As Integer: version = MSOfficeVersion
    ExcelVersion = "Unsupported"

    #If Mac Then
        If version = 14 Then
            ExcelVersion = "Mac2011"
        ElseIf version = 15 Then
            ExcelVersion = "Mac2016"
        ElseIf version = 16 Then
            ExcelVersion = "Mac2016"
        End If
    #Else
        If version = 12 Then
            ExcelVersion = "Win2007"
        ElseIf version = 14 Then
            ExcelVersion = "Win2010"
        ElseIf version = 15 Then
            ExcelVersion = "Win2013"
        ElseIf version = 16 Then
            ExcelVersion = "Win2016"
        End If
    #End If
End Function

' Returns the version of MS Office being run
'    9 = Office 2000
'   10 = Office XP / 2002
'   11 = Office 2003 & LibreOffice 3.5.2
'   12 = Office 2007
'   14 = Office 2010 or Office 2011 for Mac
'   15 = Office 2013 or Office 2016 for Mac
'   16 = Office 2016 for Mac or Windows
Public Function MSOfficeVersion() As Integer
    Dim verStr As String
    Dim startPos As Integer
    MSOfficeVersion = 0
    verStr = Application.version
    startPos = VBA.InStr(verStr, ".")
    On Error Resume Next
    If startPos > 0 Then
        MSOfficeVersion = CInt(VBA.Left(verStr, startPos - 1))
    Else
        MSOfficeVersion = CInt(verStr)
    End If
End Function

Public Function GetLicenseType() As String
    GetLicenseType = "unknown"

    #If Mac Then
        Dim homeDir As String: homeDir = "" & VBA.Replace(MacScript("return POSIX path of (path to home folder) as string"), "/Containers/com.microsoft.Excel/Data/", "")
        Dim licenseDir As String: licenseDir = homeDir & VBA.Join(VBA.Split("/Group Containers/UBF8T346G9.Office", "/"), Application.PathSeparator) & Application.PathSeparator
        If SafeDir(licenseDir & "com.microsoft.Office365.plist") <> "" Then
            GetLicenseType = "subscription"
        ElseIf SafeDir(licenseDir & "com.microsoft.e0E2OUQxNUY1LTAxOUQtNDQwNS04QkJELTAxQTI5M0JBOTk4O.plist") <> "" Then
            GetLicenseType = "subscription"
        ElseIf SafeDir(licenseDir & "e0E2OUQxNUY1LTAxOUQtNDQwNS04QkJELTAxQTI5M0JBOTk4O") <> "" Then
            GetLicenseType = "subscription"
        ElseIf SafeDir(licenseDir & "com.microsoft.Office365V2.plist") <> "" Then
            GetLicenseType = "subscription"
        ElseIf SafeDir(licenseDir & "com.microsoft.O4kTOBJ0M5ITQxATLEJkQ40SNwQDNtQUOxATL1YUNxQUO2E0e.plist") <> "" Then
            GetLicenseType = "subscription"
        ElseIf SafeDir(licenseDir & "O4kTOBJ0M5ITQxATLEJkQ40SNwQDNtQUOxATL1YUNxQUO2E0e") <> "" Then
            GetLicenseType = "subscription"
        ElseIf SafeDir(licenseDir & "Licenses", vbDirectory) <> "" Then
            Dim folderName As String, folderNum As Long, folders() As String
            folderNum = 0
            folderName = Dir(licenseDir & "Licenses" & Application.PathSeparator & "*", vbDirectory)

            ReDim Preserve folders(1 To 1)

            Do While folderName <> ""
                folderNum = folderNum + 1
                ReDim Preserve folders(1 To folderNum)
                folders(folderNum) = folderName
                folderName = Dir()
            Loop

            Dim folder
            For Each folder In folders
                Dim fileName As String
                fileName = SafeDir(licenseDir & "Licenses" & Application.PathSeparator & folder & Application.PathSeparator & "*")
                If fileName <> "" Then
                    GetLicenseType = "subscription"
                    GoTo Done
                End If
            Next
        ElseIf SafeDir("/Library/Preferences/com.microsoft.office.licensingV2.plist") <> "" Then
            GetLicenseType = "perpetual"
        End If
    #Else
        Dim oWMISrvEx       As Object   'SWbemServicesEx
        Dim oWMIObjSet      As Object   'SWbemServicesObjectSet
        Dim oWMIObjEx       As Object   'SWbemObjectEx
        Dim oWMIProp        As Object   'SWbemProperty
        Dim sWQL            As String   'WQL Statement
        Dim n               As Long     'Generic Counter

        ' Office 2010 = 59a52881-a989-479d-af46-f275c6370663
        ' Office 2013/2016 = 0ff1ce15-a989-479d-af46-f275c6370663
    On Error GoTo Win7
        sWQL = "Select Name From SoftwareLicensingProduct Where ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey <> NULL"
        Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
        Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
        GoTo DetectLicense

Win7:
    On Error GoTo Done
        sWQL = "Select Name From SoftwareLicensingProduct Where ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey <> NULL"
        Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
        Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)

DetectLicense:
        For Each oWMIObjEx In oWMIObjSet
            For Each oWMIProp In oWMIObjEx.Properties_
                If oWMIProp.name = "Name" Then
                    If VBA.InStr(VBA.LCase(oWMIProp.value), "_retail") Or VBA.InStr(VBA.LCase(oWMIProp.value), "_perp") Then
                        GetLicenseType = "perpetual"
                    ElseIf VBA.InStr(VBA.LCase(oWMIProp.value), "_sub") Then
                        GetLicenseType = "subscription"
                        GoTo Done
                    End If
                End If
            Next
        Next

        Dim licenseDir As String
        Dim fileName As String
        licenseDir = VBA.Replace(Application.UserLibraryPath, VBA.Join(VBA.Split("Roaming/Microsoft/AddIns", "/"), Application.PathSeparator), VBA.Join(VBA.Split("Local/Microsoft/Office/Licenses/5", "/"), Application.PathSeparator))
        If SafeDir(licenseDir, vbDirectory) <> "" Then
            fileName = Dir(licenseDir & Application.PathSeparator & "*")
            Do While Len(fileName) > 0
                On Error GoTo NextFile
                Dim data As String, line As String
                Dim ipt As Integer: ipt = FreeFile
                Open (licenseDir & Application.PathSeparator & fileName) For Input As ipt
                    While Not EOF(ipt)
                        Line Input #ipt, line
                        line = VBA.Trim(Application.Clean(line))
                        data = data & line
                    Wend
                Close #ipt
                Dim json
                Set json = ParseJson(data)
                Dim license
                Set license = ParseJson(Base64Decode(json.Item("License")))
                Dim ltype As String
                GetLicenseType = VBA.LCase(license.Item("LicenseType"))

                If GetLicenseType = "subscription" Then
                    GoTo Done
                End If

NextFile:
                fileName = Dir
            Loop
        End If
    #End If

Done:
End Function

Public Function IsJSCompatible() As Boolean
    IsJSCompatible = False

    #If Mac Then
        If VersionLessThan("16.24") Then Exit Function
    #Else
        If MSOfficeVersion < 16 Or Application.Build < 11601 Then Exit Function
    #End If

    If GetLicenseType <> "subscription" Then
        Exit Function
    End If

    IsJSCompatible = True
End Function
