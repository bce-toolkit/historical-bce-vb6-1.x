Attribute VB_Name = "mdlCommon"
Option Explicit

Public Sub Main()
    On Error Resume Next
    Dim sTitle As String
    If App.PrevInstance = True Then
        sTitle = App.Title
        App.Title = App.CompanyName & " : " & App.EXEName
        AppActivate sTitle
        End
    End If
    Load frmMain
    frmMain.Show
    Exit Sub
End Sub
