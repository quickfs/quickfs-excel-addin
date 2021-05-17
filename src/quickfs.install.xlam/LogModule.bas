Attribute VB_Name = "LogModule"
Option Explicit
Option Private Module

Public trigger As String

Public Sub LogMessage(msg As String)
    Dim source As String
    source = "v" & AddInVersion & " " & ThisWorkbook.name & " -"
    If VBA.Len(source) < 40 Then source = source & String(40 - VBA.Len(source), "-")
    If trigger <> "" Then msg = "(" & trigger & ") -> " & msg
    msg = "[" & VBA.Format(VBA.Now(), "yyyy-MM-dd hh:mm:ss") & "] " & source & "- " & msg
    Debug.Print (msg)
    Dim log As Integer
    log = FreeFile
    Open SavePath(AddInLogFile) For Append As log
        Print #log, msg
    Close #log
End Sub

Public Sub TrimLog(Optional days As Integer = 0)
    If days = 0 Then days = GetSetting("logRetentionDays", 3)
    Dim line As String, timestamp As String, time As Date, trimmed As Integer
    trimmed = 0
    VBA.FileCopy SavePath(AddInLogFile), SavePath(AddInLogFile & ".tmp")
    Dim out As Integer, tmp As Integer
    out = FreeFile
    Open SavePath(AddInLogFile) For Output As out
    tmp = FreeFile
    Open SavePath(AddInLogFile & ".tmp") For Input As tmp
        While Not EOF(tmp)
            Line Input #tmp, line
            line = VBA.Trim(Application.Clean(line))
            If VBA.InStr(line, "]") > 2 Then
                timestamp = VBA.Mid(line, 2, VBA.InStr(line, "]") - 2)
                time = CDate(timestamp)
                If time > VBA.Now() - days Then
                    Print #out, line
                Else
                    trimmed = trimmed + 1
                End If
            End If
        Wend
    Close #tmp
    Close #out
    
    VBA.Kill SavePath(AddInLogFile & ".tmp")
    
    If trimmed > 0 Then
        LogMessage "Trimmed " & trimmed & " messages older than " & (VBA.Now() - days)
    End If
End Sub
