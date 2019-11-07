Attribute VB_Name = "xHelpersRightClick"
Option Private Module

Public Sub RightClickMenu()
' Subroutine to generate right-click menu for
' cut / copy / paste / select all / delete
' on forms

'### Include following sub in the form module for each text box
'Private Sub textBoxName_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    If Button = vbKeyRButton Then Call RightClickMenu
'End Sub
'###

' Written by Michael Chambers

    Dim cmdBar As CommandBar
    Dim cmdButton As CommandBarButton
    
    ' Delete it if already exists
    On Error Resume Next
    CommandBars("RightClickMenu").Delete
    On Error GoTo 0
    
    ' FaceId Icons documented at http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
    Set cmdBar = CommandBars.Add(name:="RightClickMenu", Position:=msoBarPopup, temporary:=True)
    With cmdBar
        Set cmdButton = .Controls.Add(Type:=msoControlButton, temporary:=True)
        With cmdButton
            .Style = msoButtonIconAndCaption
            .FaceId = 21
            .Caption = "Cut"
            .OnAction = "cmdCut"
        End With
        Set cmdButton = .Controls.Add(Type:=msoControlButton, temporary:=True)
        With cmdButton
            .Style = msoButtonIconAndCaption
            .FaceId = 19
            .Caption = "Copy"
            .OnAction = "cmdCopy"
        End With
        Set cmdButton = .Controls.Add(Type:=msoControlButton, temporary:=True)
        With cmdButton
            .Style = msoButtonIconAndCaption
            .FaceId = 22
            .Caption = "Paste"
            .OnAction = "cmdPaste"
        End With
        Set cmdButton = .Controls.Add(Type:=msoControlButton, temporary:=True)
        With cmdButton
            .Style = msoButtonIconAndCaption
            .FaceId = 195
            .Caption = "Select All"
            .OnAction = "cmdSelectAll"
        End With
        Set cmdButton = .Controls.Add(Type:=msoControlButton, temporary:=True)
        With cmdButton
            .Style = msoButtonIconAndCaption
            .FaceId = 1786
            .Caption = "Delete"
            .OnAction = "cmdDelete"
        End With
    End With
    Application.CommandBars("RightClickMenu").ShowPopup
End Sub
Private Sub cmdCut()
' RightClickMenu Cut (Windows Ctrl-x | Mac OS/X Cmd-x)
    Dim scriptStr As String
    
    #If Mac Then
        scriptStr = "tell application " & Chr(34) & "System Events" & Chr(34) & _
                " to keystroke " & Chr(34) & "x" & Chr(34) & " using command down"
                
        MacScript (scriptStr)
    #Else
        SendKeys "^x"
    #End If
End Sub
Private Sub cmdCopy()
' RightClickMenu Copy (Windows Ctrl-c | Mac OS/X Cmd-c)
    Dim scriptStr As String
    
    #If Mac Then
        scriptStr = "tell application " & Chr(34) & "System Events" & Chr(34) & _
                " to keystroke " & Chr(34) & "c" & Chr(34) & " using command down"
                
        MacScript (scriptStr)
    #Else
        SendKeys "^c"
    #End If
End Sub
Private Sub cmdPaste()
' RightClickMenu Paste (Windows Ctrl-v | Mac OS/X Cmd-v)
    Dim scriptStr As String
    
    #If Mac Then
        scriptStr = "tell application " & Chr(34) & "System Events" & Chr(34) & _
                " to keystroke " & Chr(34) & "v" & Chr(34) & " using command down"
                
        MacScript (scriptStr)
    #Else
        SendKeys "^v"
    #End If
End Sub
Private Sub cmdSelectAll()
' RightClickMenu SelectAll (Windows Ctrl-a | Mac OS/X OSX Cmd-a)
    Dim scriptStr As String
    
    #If Mac Then
        scriptStr = "tell application " & Chr(34) & "System Events" & Chr(34) & _
                " to keystroke " & Chr(34) & "a" & Chr(34) & " using command down"
                
        MacScript (scriptStr)
    #Else
        SendKeys "^a"
    #End If
End Sub
Private Sub cmdDelete()
' RightClickMenu Delete
    Dim scriptStr As String
    
    #If Mac Then
        scriptStr = "tell application " & Chr(34) & "System Events" & Chr(34) & _
                " to key code 51"
                
        MacScript (scriptStr)
    #Else
        SendKeys "{Delete}"
    #End If
End Sub



