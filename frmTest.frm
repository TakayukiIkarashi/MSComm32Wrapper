VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows の既定値
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2820
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "設"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2460
      Width           =   495
   End
   Begin VB.CommandButton cmdReflesh 
      Caption         =   "取"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2460
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "終"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   2460
      Width           =   495
   End
   Begin VB.TextBox txtTest 
      Height          =   2295
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   0
      Top             =   60
      Width           =   4395
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "未"
      Height          =   495
      Left            =   1140
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   1
      Top             =   2460
      Width           =   495
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsConn As Boolean
Private acw As Object

Private Sub Form_Load()
  Set acw = CreateObject("MSComm32Wrapper.MSComm32")

  acw.PortOpen = True

  If (acw.PortOpen) Then
    cmdTest.Caption = "閉"
  Else
    cmdTest.Caption = "開"
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  acw.PortOpen = False
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdReflesh_Click()
  If (acw.TextData = "") Then
    txtTest.Text = "取得できず"
  Else
    txtTest.Text = acw.TextData
  End If
End Sub

Private Sub cmdSettings_Click()
  acw.ShowQRSettingsForm
End Sub

Private Sub cmdTest_Click()
  If (acw.PortOpen) Then
    acw.PortOpen = False
    cmdTest.Caption = "開"
  Else
    acw.PortOpen = True
    cmdTest.Caption = "閉"
  End If
End Sub
