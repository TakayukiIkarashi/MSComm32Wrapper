VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSComm32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------
' 変数定義
'----------------------------------------
Private MainForm As frmMSComm32         'MSComm32フォーム

'================================================================================
' TextDataプロパティ
'================================================================================
Public Property Get TextData() As String
  TextData = MainForm.TextData
End Property

Public Property Let TextData(ByVal Value As String)
  MainForm.TextData = Value
End Property

'================================================================================
' PortOpenプロパティ
'================================================================================
Public Property Get PortOpen() As Boolean
  PortOpen = MainForm.MSComm1.PortOpen
End Property

Public Property Let PortOpen(ByVal Value As Boolean)
  If (Value) Then
    Call OpenComPort
  Else
    Call CloseComPort
  End If
End Property

'================================================================================
' コンストラクタ
'================================================================================
Private Sub Class_Initialize()
  '設定ファイルからRS232C接続の設定を取得します
  Call GetIniFile

  'メインフォームをロードします
  Set MainForm = New frmMSComm32
  MainForm.Visible = False
End Sub

'================================================================================
' デストラクタ
'================================================================================
Private Sub Class_Terminate()
  'COMポートを閉じます
  Call CloseComPort

  'メインフォームを閉じます。
  Unload MainForm
End Sub

'================================================================================
' 概要　：「RS232C設定」フォームを起動します
' 引数　：なし
' 戻り値：なし
'================================================================================
Public Sub ShowQRSettingsForm()

  '「QR設定」ウィンドウを定義します
  Dim f As New frmQRSettings
  f.QREnabled = ComInfo.QREnabled
  f.portNo = ComInfo.portNo
  f.speed = ComInfo.speed
  f.dataBits = ComInfo.dataBits
  f.parity = ComInfo.parity
  f.stopBits = ComInfo.stopBits

  '「QR設定」ウィンドウを表示します
  f.Show vbModal

  '「QR設定」ウィンドウの内容を構造体に反映します
  ComInfo.QREnabled = f.QREnabled
  ComInfo.portNo = f.portNo
  ComInfo.speed = f.speed
  ComInfo.dataBits = f.dataBits
  ComInfo.parity = f.parity
  ComInfo.stopBits = f.stopBits

  '「QR設定」ウィンドウを閉じます
  Unload f

  'RS232Cの設定を設定ファイルに保存します。
  Call SetIniFile

End Sub

'================================================================================
' 概要　：COMポートを開きます
' 引数　：なし
' 戻り値：なし
'================================================================================
Private Sub OpenComPort()

  'すでにCOMポートが開いている場合は処理を抜けます
  If (MainForm.MSComm1.PortOpen) Then
    Exit Sub
  End If

  On Error GoTo Exception

  With MainForm.MSComm1
    'ポート番号の設定
    .CommPort = ComInfo.portNo

    'COMポート設定文字列を格納する変数を定義します
    Dim s As String
    s = ""
    s = s & CStr(ComInfo.speed) & ","
    s = s & ComInfo.parity & ","
    s = s & CStr(ComInfo.dataBits) & ","
    s = s & CStr(ComInfo.stopBits)

    'COMポート設定文字列
    .Settings = s

    'Inputプロパティ使用時にバッファ全体を読み取るように設定します
    .InputLen = 0

    'ポートを開きます
    .PortOpen = True
    .RTSEnable = True
  End With

  If (MainForm.MSComm1.PortOpen = False) Then
    Call Err.Raise(ErrNumber.CannotPortOpen, "OpenComPort", "ポートを開けません。")
  End If

  Exit Sub

Exception:
  Call Err.Raise(Err.Number, "OpenComPort", Err.Description)

End Sub

'================================================================================
' 概要　：COMポートを閉じます
' 引数　：なし
' 戻り値：なし
'================================================================================
Private Sub CloseComPort()

  'すでにCOMポートが閉じていれば処理を抜けます
  If (MainForm.MSComm1.PortOpen = False) Then
    Exit Sub
  End If

  On Error GoTo Exception

  'ポートを閉じます
  MainForm.MSComm1.RTSEnable = False
  MainForm.MSComm1.PortOpen = False

  Exit Sub

Exception:
  Call Err.Raise(Err.Number, "OpenComPort", Err.Description)

End Sub
