Attribute VB_Name = "QuickfsUnlinkModule"
Option Explicit
Option Private Module

Public Sub UnlinkFormulas()
    On Error GoTo ShowWarning
    
    If Not ActiveWorkbook.Saved Then
        MsgBox _
            Title:="[QuickFS] Unlink Canceled", _
            Prompt:="This workbook contains unsaved changes. You must save before it can be unlinked.", _
            Buttons:=vbExclamation
        Exit Sub
    End If
    
    Dim wbName As String
    wbName = ActiveWorkbook.name
    wbName = Replace(wbName, ".xlsm", "")
    wbName = Replace(wbName, ".xlsx", "")
    wbName = Replace(wbName, ".xls", "")

    Dim choice As Variant
    choice = MsgBox( _
        Title:="[QuickFS] Unlink Confirmation", _
        Prompt:="This will save a copy of the current workbook with all QFS formulas replaced by their current values. Do you wish to continue?", _
        Buttons:=vbYesNo Or vbQuestion)
    Select Case choice
        Case vbYes
            Dim fileSaveName As Variant
            #If Mac Then
                fileSaveName = Application.GetSaveAsFilename( _
                    InitialFileName:=wbName & " - unlinked")
            #Else
                fileSaveName = Application.GetSaveAsFilename( _
                    InitialFileName:=wbName & " - unlinked", _
                    fileFilter:="Excel Workbook (*.xlsx), *.xlsx")
            #End If
    
            If TypeName(fileSaveName) <> "Boolean" Then
                Application.DisplayAlerts = False
            
                Dim calcType: calcType = Application.Calculation
                Application.Calculation = xlCalculationManual
                Dim r As Range, i As Long
                For i = 1 To Sheets.count
                    On Error Resume Next
                    For Each r In Sheets(i).UsedRange.SpecialCells(xlCellTypeFormulas)
                        If r.Formula Like "*QFS*" Then r.value = r.value
                    Next r
                Next i
                Application.Calculation = calcType
                
                ActiveWorkbook.SaveAs Filename:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
                Application.DisplayAlerts = True
            End If
    End Select
    Exit Sub
    
ShowWarning:
    MsgBox _
        Title:="[QuickFS] Unlink Error", _
        Prompt:="This workbook cannot be unlinked. Please contact support@quickfs.net if this problem persists.", _
        Buttons:=vbCritical
End Sub


