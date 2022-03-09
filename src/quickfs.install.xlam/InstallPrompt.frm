VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InstallPrompt 
   Caption         =   "QuickFS Installer"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7815
   OleObjectBlob   =   "InstallPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InstallPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private uninstallEnabled As Boolean

Private Sub UserForm_Initialize()
    Me.VersionLabel.Caption = "Excel Add-In v" & AddInVersion
    
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If (VBA.LCase(i.name) = VBA.LCase(AddInInstalledFile)) And VBA.LCase(i.FullName) = VBA.LCase(SavePath(i.name)) Then
            Set installed = i
            Exit For
        End If
    Next i
    
    Dim msg As String, UpgradeVersion As String, CurrentVersion As String
    UpgradeVersion = AddInVersion(ThisWorkbook.name)
    If Not installed Is Nothing Then
        CurrentVersion = AddInVersion(installed.name)
        If CurrentVersion = "" Then
            uninstallEnabled = False
            Me.InstallButton.Caption = "Install"
            Me.UninstallButton.BackColor = RGB(200, 200, 200)
            Me.UninstallButtonBg.BackColor = RGB(200, 200, 200)
        Else
            uninstallEnabled = True
            Me.InstallButton.Caption = "Update"
            Me.UninstallButton.BackColor = RGB(230, 33, 23)
            Me.UninstallButtonBg.BackColor = RGB(230, 33, 23)
        End If
    Else
        uninstallEnabled = False
        Me.InstallButton.Caption = "Install"
        Me.UninstallButton.BackColor = RGB(200, 200, 200)
        Me.UninstallButtonBg.BackColor = RGB(200, 200, 200)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        CancelInstall
    End If
End Sub

Private Sub InstallButton_Click()
    InstallButtonBg_Click
End Sub

Private Sub InstallButtonBg_Click()
    If installing Then Exit Sub
    Me.InstallButton.BackColor = RGB(10, 37, 88)
    Me.InstallButtonBg.BackColor = RGB(10, 37, 88)
    
    ' Note! Because of a bug in Mac2011, this form
    ' needs to be unloaded or workbooks will fail to close
    ' during the install function.
    ' See https://stackoverflow.com/questions/10612502/error-when-closing-an-opened-workbook-in-vba-userform
    If ExcelVersion = "Mac2011" Then Unload Me
    
    FinishInstalling
End Sub

Private Sub InstallButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.InstallButton.BackColor = RGB(10, 37, 88)
    Me.InstallButtonBg.BackColor = RGB(10, 37, 88)
End Sub

Private Sub InstallButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.InstallButton.BackColor = RGB(21, 81, 195)
    Me.InstallButtonBg.BackColor = RGB(21, 81, 195)
End Sub

Private Sub UninstallButton_Click()
    UninstallButtonBg_Click
End Sub

Private Sub UninstallButtonBg_Click()
    If uninstalling Or Not uninstallEnabled Then Exit Sub
    Me.UninstallButton.BackColor = RGB(168, 22, 15)
    Me.UninstallButtonBg.BackColor = RGB(168, 22, 15)

    ' Note! Because of a bug in Mac2011, this form
    ' needs to be unloaded or workbooks will fail to close
    ' during the uninstall function.
    ' See https://stackoverflow.com/questions/10612502/error-when-closing-an-opened-workbook-in-vba-userform
    If ExcelVersion = "Mac2011" Then Unload Me
    
    UninstallAddIn
End Sub

Private Sub UninstallButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not uninstallEnabled Then Exit Sub
    Me.UninstallButton.BackColor = RGB(168, 22, 15)
    Me.UninstallButtonBg.BackColor = RGB(168, 22, 15)
End Sub

Private Sub UninstallButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not uninstallEnabled Then Exit Sub
    Me.UninstallButton.BackColor = RGB(230, 33, 23)
    Me.UninstallButtonBg.BackColor = RGB(230, 33, 23)
End Sub

