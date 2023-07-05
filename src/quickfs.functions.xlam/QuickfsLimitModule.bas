Attribute VB_Name = "QuickfsLimitModule"
Option Explicit
Option Private Module

Private RedisplayWarning As Date

Public Sub ShowRateLimitWarning(Optional reset As Boolean = True)
    MsgBox _
        Title:="[QuickFS] Limit Exceeded", _
        Prompt:="You have exhausted your QuickFS data limit. Try again later or contact support@quickfs.net to request a limit increase.", _
        Buttons:=vbCritical
    If reset Then SetRateLimitTimer
End Sub

Public Sub ShowRequestTimeoutMessage()
    MsgBox _
        Title:="[QuickFS] Request Timeout Error", _
        Prompt:="Our server took too long to respond, which caused your request to fail. Please email your spreadsheet to support@quickfs.net so that we can fix the issue.", _
        Buttons:=vbCritical
End Sub

Public Function IsRateLimited()
    If RedisplayWarning > Now() Then
        IsRateLimited = True
    Else
        IsRateLimited = False
    End If
End Function

Public Sub ClearRateLimit()
    RedisplayWarning = Now() - 1
End Sub

Private Sub SetRateLimitTimer()
    RedisplayWarning = Now() + (5 / (60 * 24))
End Sub

