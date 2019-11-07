VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mac2016CredentialsForm 
   Caption         =   "QuickFS Login"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   OleObjectBlob   =   "Mac2016CredentialsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mac2016CredentialsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private password As String

Private Sub IssueDetailLink_Click()
    ThisWorkbook.FollowHyperlink "https://appletoolbox.com/2015/10/mouse-cursor-pointer-disappears-invisible-missing-fix/"
End Sub


Private Sub passBox_Change()
    Dim text As String, i As Integer, oldLen As Integer, newLen As Integer, diff As Integer, chars As String
    text = Me.passBox.value
    oldLen = VBA.Len(password)
    newLen = VBA.Len(text)
    diff = newLen - oldLen
    If diff > 0 Then
        chars = VBA.Right(text, diff)
        password = password & chars
        text = ""
        For i = 1 To newLen
            text = text & "*"
        Next i
        Me.passBox.value = text
    ElseIf diff < 0 Then
        password = ""
        Me.passBox.value = ""
    End If
End Sub

Private Sub passBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        Me.passBox.SelStart = 0
        Me.passBox.SelLength = VBA.Len(Me.passBox.value)
    End If
End Sub

Private Sub emailBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And password <> "" And Me.emailBox.value <> "" Then
        LoginButton_Click
    End If
End Sub

Private Sub passBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.passBox.SelStart = 0
    Me.passBox.SelLength = VBA.Len(Me.passBox.value)
End Sub

Private Sub SignUpLabel_Click()
    ThisWorkbook.FollowHyperlink SIGNUP_URL
End Sub

Private Sub UserForm_Initialize()
    Me.emailBox.SetFocus
End Sub

Private Sub LoginButton_Click()
    Dim Success As Boolean
    Success = Login(Me.emailBox.value, password)
    If Success Then
        Unload Me
        Application.CalculateFull
    End If
End Sub

Private Sub LoginButtonBg_Click()
    Dim Success As Boolean
    Success = Login(Me.emailBox.value, password)
    If Success Then
        Unload Me
        Application.CalculateFull
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

