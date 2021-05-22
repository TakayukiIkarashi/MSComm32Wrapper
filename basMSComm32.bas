Attribute VB_Name = "basMSComm32"
Option Explicit

'----------------------------------------
' 構造体定義
'----------------------------------------
'QRコードリーダーの状態を格納する構造体
Public Type ComTable
  QREnabled As Boolean                  'QRコード使用可否
  portNo As Integer                     'ポート番号
  speed As Long                         '通信速度
  parity As String                      'パリティ
  dataBits As Integer                   'データビット
  stopBits As Single                    'ストップビット
End Type
Public ComInfo As ComTable

'プロジェクト内で用意したエラー
Public Enum ErrNumber
  CannotPortOpen = 513                  'ポートが開けない
  UnknownEvent = 514                    '認識できないイベントエラー
End Enum

'----------------------------------------
' INIファイル定義
'----------------------------------------
Public Const INIFILE_NAME_ME As String = "mscomm32wrapper_rs232c.ini" 'INIファイル名
Public Const INISEC_RS232C As String = "RS232C"                       'セクション：RS232C
Public Const INIKEY_RS232C_ENABLED As String = "ENABLED"              '　QRコードリーダー使用可否
Public Const INIKEY_RS232C_PORTNO As String = "PORTNO"                '　ポート番号
Public Const INIKEY_RS232C_SPEED As String = "SPEED"                  '　通信速度
Public Const INIKEY_RS232C_DATABITS As String = "DATABITS"            '　データビット
Public Const INIKEY_RS232C_PARITY As String = "PARITY"                '　パリティ
Public Const INIKEY_RS232C_STOPBITS As String = "STOPBITS"            '　ストップビット

'================================================================================
' 概要　：このアプリケーションで使用するINIファイルを読み込み、
' 　　　　前回の起動状態を復元します
' 引数　：なし
' 戻り値：なし
'================================================================================
Public Sub GetIniFile()

  '[QRコードリーダーを使う]
  Dim qrEn As Boolean
  qrEn = CBool(getINIValue(INIKEY_RS232C_ENABLED, INISEC_RS232C, "0", App.Path & "\" & INIFILE_NAME_ME))

  '[ポート番号]
  Dim portNo As Integer
  portNo = Val(getINIValue(INIKEY_RS232C_PORTNO, INISEC_RS232C, "", App.Path & "\" & INIFILE_NAME_ME))
  If (portNo = 0) Then
    portNo = 1
  End If

  '[通信速度]
  Dim speed As Long
  speed = Val(getINIValue(INIKEY_RS232C_SPEED, INISEC_RS232C, "", App.Path & "\" & INIFILE_NAME_ME))
  If (speed = 0) Then
    speed = 9600
  End If

  '[パリティ]
  Dim parity As String
  parity = getINIValue(INIKEY_RS232C_PARITY, INISEC_RS232C, "", App.Path & "\" & INIFILE_NAME_ME)
  If (parity = "") Then
    parity = "N"
  End If

  '[データビット]
  Dim dataBits As Integer
  dataBits = Val(getINIValue(INIKEY_RS232C_DATABITS, INISEC_RS232C, "8", App.Path & "\" & INIFILE_NAME_ME))

  '[ストップビット]
  Dim stopBits As Single
  stopBits = Val(getINIValue(INIKEY_RS232C_STOPBITS, INISEC_RS232C, "1", App.Path & "\" & INIFILE_NAME_ME))

  '構造体にRS232C情報をセットします
  With ComInfo
    .QREnabled = qrEn
    .portNo = portNo
    .speed = speed
    .parity = parity
    .dataBits = dataBits
    .stopBits = stopBits
  End With

End Sub

'================================================================================
' 概要　：このアプリケーションで使用するINIファイルにRS232C設定情報を書き込みます
' 引数　：なし
' 戻り値：なし
'================================================================================
Public Sub SetIniFile()

  '[QRコードリーダーを使う]
  Call setINIValue(ComInfo.QREnabled, INIKEY_RS232C_ENABLED, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[ポート番号]
  Call setINIValue(ComInfo.portNo, INIKEY_RS232C_PORTNO, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[通信速度]
  Call setINIValue(ComInfo.speed, INIKEY_RS232C_SPEED, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[データビット]
  Call setINIValue(ComInfo.dataBits, INIKEY_RS232C_DATABITS, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[パリティ]
  Call setINIValue(ComInfo.parity, INIKEY_RS232C_PARITY, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[ストップビット]
  Call setINIValue(ComInfo.stopBits, INIKEY_RS232C_STOPBITS, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

End Sub
