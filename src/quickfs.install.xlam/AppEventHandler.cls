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

Private Sub App_AfterCalculate()
    If Not GetSetting("updateOnAfterCalculate", True) Then Exit Sub
    trigger = "AfterCalculate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal Wn As Window)
    If Not GetSetting("updateOnWindowActivate", True) Then Exit Sub
    trigger = "WindowActivate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WindowDeactivate(ByVal wb As Workbook, ByVal Wn As Window)
    If Not GetSetting("updateOnWindowDeactivate", True) Then Exit Sub
    trigger = "WindowDeactivate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookActivate(ByVal wb As Workbook)
    If Not GetSetting("updateOnWorkbookActivate", True) Then Exit Sub
    trigger = "WorkbookActivate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookDeactivate(ByVal wb As Workbook)
    If Not GetSetting("updateOnWorkbookDeactivate", True) Then Exit Sub
    trigger = "WorkbookDeactivate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    If Not GetSetting("updateOnWorkbookOpen", True) Then Exit Sub
    trigger = "WorkbookOpen"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    If Not GetSetting("updateOnWorkbookBeforeClose", True) Then Exit Sub
    trigger = "WorkbookBeforeClose"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookAfterSave(ByVal wb As Workbook, ByVal Success As Boolean)
    If Not GetSetting("updateOnWorkbookAfterSave", True) Then Exit Sub
    trigger = "WorkbookAfterSave"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_NewWorkbook(ByVal wb As Workbook)
    If Not GetSetting("updateOnNewWorkbook", True) Then Exit Sub
    trigger = "NewWorkbook"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    If Not GetSetting("updateOnSheetActivate", True) Then Exit Sub
    trigger = "SheetActivate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    If Not GetSetting("updateOnSheetCalculate", True) Then Exit Sub
    trigger = "SheetCalculate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    If Not GetSetting("updateOnSheetDeactivate", True) Then Exit Sub
    trigger = "SheetDeactivate"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not GetSetting("updateOnSheetSelectionChange", True) Then Exit Sub
    trigger = "SheetSelectionChange"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub

Private Sub App_WorkbookNewSheet(ByVal wb As Workbook, ByVal Sh As Object)
    If Not GetSetting("updateOnWorkbookNewSheet", True) Then Exit Sub
    trigger = "WorkbookNewSheet"
    CloseInstaller
    AutoUpdateCheck
    PromoteStagedUpdate
    trigger = ""
End Sub
