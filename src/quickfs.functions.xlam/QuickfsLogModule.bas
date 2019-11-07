Attribute VB_Name = "QuickfsLogModule"
Option Explicit
Option Private Module

Public trigger As String

Public Sub ShowMessages()
    Application.Run (AddInManagerFile & "!TrimLog")
    #If Mac Then
        Dim file As String
        file = LocalPath(AddInLogFile)
        
        ' Remove the disk specifier (e.g. 'Mac HD:')
        If Application.PathSeparator = ":" Then file = VBA.Mid(file, VBA.InStr(file, ":") + 1)
        
        ' Normalize the path separator
        file = VBA.Replace(file, Application.PathSeparator, "/")
        
        Call xHelpersWeb.ExecuteInShell("open '" & file & "'")
    #Else
        Dim shell As Object
        Set shell = CreateObject("Shell.Application")
        shell.Open LocalPath(AddInLogFile)
    #End If
End Sub

Public Sub LogMessage(ByVal msg As String, Optional ByVal key As String = "")
    Dim source As String
    source = "v" & AddInVersion & " " & ThisWorkbook.name & " -"
    If VBA.Len(source) < 40 Then source = source & String(40 - VBA.Len(source), "-")
    If trigger <> "" Then msg = "(" & trigger & ") -> " & msg
    If key <> "" Then msg = "(" & key & ") -> " & msg
    msg = "[" & VBA.Format(VBA.Now(), "yyyy-MM-dd hh:mm:ss") & "] " & source & "- " & msg
    Debug.Print msg
    Dim log As Integer
    log = FreeFile
    Open LocalPath(AddInLogFile) For Append As log
        Print #log, msg
    Close #log
End Sub
