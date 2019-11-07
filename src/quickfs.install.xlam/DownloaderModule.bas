Attribute VB_Name = "DownloaderModule"
Option Explicit
Option Private Module

#If Mac Then

Public Sub DownloadFile(url As String, file As String)
    Dim result As ShellResult
    If Application.PathSeparator = ":" Then file = VBA.Mid(file, VBA.InStr(file, ":") + 1)
    file = VBA.Replace(file, Application.PathSeparator, "/")
    result = xHelpersWeb.ExecuteInShell("curl -L -s -o '" & file & "' " & url)
    If result.ExitCode > 0 Then xHelpersWeb.RaiseCurlError result, url
End Sub

#ElseIf VBA7 Then

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Sub DownloadFile(url As String, file As String)
    Dim result As Long
    result = URLDownloadToFile(0, url, file, 0, 0)
End Sub

#Else

Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Sub DownloadFile(url As String, file As String)
    Dim result As Long
    result = URLDownloadToFile(0, url, file, 0, 0)
End Sub

#End If

