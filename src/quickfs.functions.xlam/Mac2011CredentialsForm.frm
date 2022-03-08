VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mac2011CredentialsForm 
   Caption         =   "QuickFS Login"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   -525
   ClientWidth     =   6060
   OleObjectBlob   =   "Mac2011CredentialsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mac2011CredentialsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub emailBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Me.passBox.value <> "" And Me.emailBox.value <> "" Then
        LoginButton_Click
    End If
End Sub

Private Sub SignUpLabel_Click()
    ThisWorkbook.FollowHyperlink SIGNUP_URL
End Sub

Private Sub UserForm_Initialize()
    Me.emailBox.SetFocus
End Sub

Private Sub emailBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = vbKeyRButton Then Call RightClickMenu
End Sub

Private Sub passBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = vbKeyRButton Then Call RightClickMenu
End Sub

Private Sub LoginButton_Click()
    Dim Success As Boolean
    Success = Login(Me.emailBox.value, Me.passBox.value)
    If Success Then
        Unload Me
    End If
End Sub

Private Sub LoginButtonBg_Click()
    Dim Success As Boolean
    Success = Login(Me.emailBox.value, Me.passBox.value)
    If Success Then
        Unload Me
    End If
End Sub

Private Sub LoginButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(10, 37, 88)
    Me.LoginButtonBg.BackColor = RGB(10, 37, 88)
End Sub

Private Sub LoginButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(21, 81, 195)
    Me.LoginButtonBg.BackColor = RGB(21, 81, 195)
End Sub

Private Sub LoginButtonBg_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(10, 37, 88)
    Me.LoginButtonBg.BackColor = RGB(10, 37, 88)
End Sub

Private Sub LoginButtonBg_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(21, 81, 195)
    Me.LoginButtonBg.BackColor = RGB(21, 81, 195)
End Sub
