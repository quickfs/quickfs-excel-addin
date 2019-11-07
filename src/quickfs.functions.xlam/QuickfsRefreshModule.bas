Attribute VB_Name = "QuickfsRefreshModule"
Option Explicit
Option Private Module

Public Sub RefreshData()
    On Error GoTo EnableCache
    
    If IsRateLimited Then
        ShowRateLimitWarning
        Exit Sub
    End If
    
    StartRecache
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Calculate
    Next
    
EnableCache:
    StopRecache
End Sub

