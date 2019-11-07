VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIKeyHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private activeKey As String
Private hasReadKey As Boolean

Private Sub Class_Initialize()
    activeKey = ""
    hasReadKey = False
End Sub

Public Sub StoreApiKey(key As String)
    activeKey = key
    hasReadKey = True
    LogMessage "Using API key " & activeKey
    WriteKey activeKey
    InvalidateAppRibbon
End Sub

Public Function GetAPIKey()
    If Not hasReadKey Then
        activeKey = ReadKey
        hasReadKey = True
        If activeKey <> "" Then LogMessage "Using API key " & activeKey
    End If
    GetAPIKey = activeKey
End Function

Public Sub ClearAPIKey()
    activeKey = ""
    hasReadKey = True
    LogMessage "Logging out"
    WriteKey activeKey
    InvalidateAppRibbon
End Sub

Private Sub WriteKey(key As String)
    Dim keyfile As String
    keyfile = LocalPath(AddInKeyFile)
    Dim file As Integer
    file = FreeFile
    Open keyfile For Output As file
    Write #file, key
    Close #file
End Sub

Private Function ReadKey()
    On Error GoTo Finish
    Dim key As String, keyfile As String
    key = ""
    keyfile = LocalPath(AddInKeyFile)
    Dim ipt As Integer
    ipt = FreeFile
    Open keyfile For Input As ipt
    Line Input #ipt, key
    key = VBA.Replace(key, """", "")
    Close #ipt
Finish:
    ReadKey = key
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Old method of key storage is commented out below.
' Didn't really work because it relied on saving the key
' in the add-on workbook, which was unreliable because
' saving sometimes fails without explanation, and in some
' environments the workbook is read-only anyways.
' Keeping the code here as reference because it shows how
' to use application events, which will probably be useful
' at some point in the future.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private WithEvents App As Application
'Private keyToStore As String
'Private isStoring As Boolean

'Private Sub Class_Initialize()
'    Set App = Application
'    keyToStore = ""
'    isStoring = False
'End Sub

'Public Sub StoreApiKey(key As String)
    'keyToStore = key
    'StoreKey
'End Sub

'Public Sub ClearAPIKey()
    'keyToStore = ""
    'StoreKey
'End Sub

'Public Function GetAPIKey()
    'If keyToStore <> "" Then
    '    GetAPIKey = keyToStore
    'Else
    '    GetAPIKey = VBA.Trim(ThisWorkbook.Sheets("API_Credentials").Range("APIkey").value)
    'End If
'End Function

'Private Sub App_SheetChange(ByVal Sh As Object, ByVal Source As Range)
'    If isStoring = False And Not keyToStore = "" Then
'        StoreKey
'    End If
'End Sub

'Private Sub StoreKey()
'    On Error GoTo FinishStoring
'    isStoring = True
'    ThisWorkbook.Sheets("API_Credentials").Range("APIkey").value = keyToStore
'    Application.DisplayAlerts = False
'    ThisWorkbook.Save
'    keyToStore = ""
'FinishStoring:
'    Application.DisplayAlerts = True
'    isStoring = False
'    InvalidateAppRibbon
'End Sub


