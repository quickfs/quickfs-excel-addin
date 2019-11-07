Attribute VB_Name = "QuickfsHelpModule"
Option Explicit
Option Private Module

Public Function LoadHelp()
    ThisWorkbook.FollowHyperlink HELP_URL
End Function


