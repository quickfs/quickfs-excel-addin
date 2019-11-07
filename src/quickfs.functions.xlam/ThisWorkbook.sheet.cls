Private handler As AppEventHandler

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error GoTo Finish
    If SafeDir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> "" Then
'        Dim openName As String
'        openName = Workbooks("vbaDeveloper.xlam").name
'        Call Application.Run("vbaDeveloper.xlam!exportVbProject", "quickfsFunctions")
    End If
Finish:
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RemoveCustomMenu
    LogMessage "Unloaded add-in functions"
End Sub

Private Sub Workbook_Open()
    LogMessage "Loading add-in functions"
    
    Set handler = New AppEventHandler
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call AddUDFCategoryDescription
    Call RemoveCustomMenu
    Call UpdateCustomMenu
    
    CheckQuota blockEvents:=True
    
    LogMessage "Add-in functions are ready"
End Sub