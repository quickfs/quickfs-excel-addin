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
