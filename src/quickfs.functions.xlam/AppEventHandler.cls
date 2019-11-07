VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    trigger = "WorkbookOpen"
    QuickfsFixLinksModule.FixAddInLinks wb
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_AfterCalculate()
    trigger = "AfterCalculate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal Wn As Window)
    trigger = "WindowActivate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WindowDeactivate(ByVal wb As Workbook, ByVal Wn As Window)
    trigger = "WindowDeactivate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookActivate(ByVal wb As Workbook)
    Call AddUDFCategoryDescription
    trigger = "WorkbookActivate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookDeactivate(ByVal wb As Workbook)
    trigger = "WorkbookDeactivate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    trigger = "WorkbookBeforeClose"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookAfterSave(ByVal wb As Workbook, ByVal Success As Boolean)
    trigger = "WorkbookAfterSave"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_NewWorkbook(ByVal wb As Workbook)
    trigger = "NewWorkbook"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    trigger = "SheetActivate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    trigger = "SheetCalculate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    trigger = "SheetDeactivate"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    trigger = "SheetSelectionChange"
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookNewSheet(ByVal wb As Workbook, ByVal Sh As Object)
    trigger = "WorkbookNewSheet"
    PromoteStagedUpdate
    trigger = ""
End Sub

