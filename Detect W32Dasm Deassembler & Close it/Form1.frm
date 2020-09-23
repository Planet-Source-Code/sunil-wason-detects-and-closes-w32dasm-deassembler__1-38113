VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFound 
      Caption         =   "Raise system error if W32 Dasm Deassembler found"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Timer1_Timer()

Dim dummyval As Long
Dim HwndDasm As Long

HwndDasm = WinDasmHwnd
'If W32Dasm found
If HwndDasm <> 0 Then
    'Close W32Dasm
    dummyval = sendmessagebystring(HwndDasm, WM_CLOSE, 0, 0)
    If chkFound = 1 Then
        'Raise an exception
        RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
    End If
End If

End Sub
