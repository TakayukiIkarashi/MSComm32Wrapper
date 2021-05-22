VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMSComm32 
   Caption         =   "Form1"
   ClientHeight    =   555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   1800
   StartUpPosition =   3  'Windows の既定値
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
End
Attribute VB_Name = "frmMSComm32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' 変数定義
'----------------------------------------
Private mstrTextData As String          'TextDataプロパティ

'================================================================================
' TextDataプロパティ
'================================================================================
Public Property Get TextData() As String
  TextData = mstrTextData
End Property

Public Property Let TextData(ByVal Value As String)
  mstrTextData = Value
End Property

'================================================================================
' MSComm32 OnCommイベント
'================================================================================
Private Static Sub MSComm1_OnComm()

  'COMイベントを取得します
  Select Case MSComm1.CommEvent

    '受信イベント
    Case comEvReceive
      Dim b() As Byte
      Dim s() As Byte
      b() = MSComm1.Input
      s() = CStr(s()) & CStr(b())

      If (s(UBound(s)) = &HD) Then
        Dim buffer As String
        buffer = CStr(s())
        Erase s

        TextData = StrConv(buffer, vbUnicode)
      End If

    'その他のイベント
    Case comEvSend, comEvCTS, comEvDSR, comEvCD, comEvRing, comEvEOF

    '認識できないイベント
    Case Else
      Call Err.Raise(ErrNumber.UnknownEvent, "MSComm1_OnComm", "認識できないイベントです。")
      Exit Sub

  End Select

End Sub
