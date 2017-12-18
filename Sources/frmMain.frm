VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "*"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "*"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdBalance 
      Caption         =   "*"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtOutput 
      Height          =   330
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   4575
   End
   Begin VB.TextBox txtOutput 
      Height          =   330
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtInput 
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label lblMessages 
      Caption         =   "*"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblMessages 
      Caption         =   "*"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBalance_Click()
    On Error Resume Next
    Dim strEquation As String
    Dim strOut As String
    Dim lpCurrent As Long
    Dim bpResult As Boolean
    Dim rsBalanced() As Long
    Dim psSides As New Collection
    Dim psSide1 As New Collection, psSide2 As New Collection
    strEquation = txtInput.Text
    strEquation = RemoveSpace(strEquation)
    strEquation = ResolveStringCC(strEquation, "==", "=", "++", "+", "`", "", ".", "")
    If Trim(strEquation) = vbNullString Then
        MsgBox LoadResString(106), vbExclamation, LoadResString(32)
        txtInput.SetFocus
        Exit Sub
    End If
    bpResult = BalanceCE(strEquation, rsBalanced())
    If bpResult = False Then
        MsgBox LoadResString(107), vbCritical, LoadResString(16)
        txtInput.SetFocus
        Exit Sub
    End If
    ClearCollection psSides
    ClearCollection psSide1
    ClearCollection psSide2
    ResolveCommandEX Trim(strEquation), psSides, "="
    ResolveCommandEX psSides.Item(1), psSide1, "+"
    ResolveCommandEX psSides.Item(2), psSide2, "+"
    strOut = vbNullString
    For lpCurrent = 1 To psSide1.Count
        strOut = strOut & IIf(rsBalanced(lpCurrent) <> 1, Trim(Str(rsBalanced(lpCurrent))), "") & psSide1.Item(lpCurrent) & IIf(lpCurrent = psSide1.Count, "=", "+")
    Next lpCurrent
    For lpCurrent = 1 To psSide2.Count
        strOut = strOut & IIf(rsBalanced(psSide1.Count + lpCurrent) <> 1, Trim(Str(rsBalanced(psSide1.Count + lpCurrent))), "") & psSide2.Item(lpCurrent) & IIf(lpCurrent = psSide2.Count, "", "+")
    Next lpCurrent
    txtOutput(0).Text = strOut
    strOut = vbNullString
    For lpCurrent = 1 To UBound(rsBalanced())
        strOut = strOut & Trim(Str(rsBalanced(lpCurrent))) & IIf(lpCurrent = UBound(rsBalanced()), "", ",")
    Next lpCurrent
    txtOutput(1).Text = strOut
End Sub

Private Sub cmdExit_Click()
    On Error Resume Next
    Unload frmMain
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strInputRG As String
    frmMain.Caption = LoadResString(101)
    lblMessages(0).Caption = LoadResString(102)
    lblMessages(1).Caption = LoadResString(103)
    cmdBalance.Caption = LoadResString(104)
    cmdExit.Caption = LoadResString(105)
    strInputRG = GetSetting(App.EXEName, "Windows", "Top", Trim(Str(Screen.Height / 2 - frmMain.ScaleHeight / 2)))
    If IsNumeric(strInputRG) = True Then
        frmMain.Top = Val(strInputRG)
    End If
    strInputRG = GetSetting(App.EXEName, "Windows", "Left", Trim(Str(Screen.Width / 2 - frmMain.ScaleWidth / 2)))
    If IsNumeric(strInputRG) = True Then
        frmMain.Left = Val(strInputRG)
    End If
    txtInput.Text = GetSetting(App.EXEName, "Default", "Equation", LoadResString(201))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    With frmMain
        SaveSetting App.EXEName, "Windows", "Top", Trim(Str(.Top))
        SaveSetting App.EXEName, "Windows", "Left", Trim(Str(.Left))
    End With
    SaveSetting App.EXEName, "Default", "Equation", txtInput.Text
    Cancel = False
    DoEvents
    End
End Sub
