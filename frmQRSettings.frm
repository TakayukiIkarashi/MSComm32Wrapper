VERSION 5.00
Begin VB.Form frmQRSettings 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "QR設定"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "メイリオ"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chkQREnabled 
      Caption         =   "QRコードリーダーを使う"
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   3195
   End
   Begin VB.Frame fraEnabled 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4875
      Begin VB.ComboBox cboSpeed 
         Height          =   480
         ItemData        =   "frmQRSettings.frx":0000
         Left            =   2700
         List            =   "frmQRSettings.frx":006D
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   5
         Top             =   1080
         Width           =   1620
      End
      Begin VB.ComboBox cboPortNo 
         Height          =   480
         ItemData        =   "frmQRSettings.frx":00DA
         Left            =   2700
         List            =   "frmQRSettings.frx":0115
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   3
         Top             =   600
         Width           =   1620
      End
      Begin VB.Frame faConnection 
         Caption         =   "接続設定"
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   2190
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   4380
         Begin VB.ComboBox cboDataBits 
            Height          =   480
            ItemData        =   "frmQRSettings.frx":0180
            Left            =   2460
            List            =   "frmQRSettings.frx":0193
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   8
            Top             =   480
            Width           =   1620
         End
         Begin VB.ComboBox cboParity 
            Height          =   480
            ItemData        =   "frmQRSettings.frx":01A6
            Left            =   2460
            List            =   "frmQRSettings.frx":01B9
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   10
            Top             =   960
            Width           =   1620
         End
         Begin VB.ComboBox cboStopBits 
            Height          =   480
            ItemData        =   "frmQRSettings.frx":01DB
            Left            =   2460
            List            =   "frmQRSettings.frx":01E8
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   12
            Top             =   1440
            Width           =   1620
         End
         Begin VB.Label lbRs 
            AutoSize        =   -1  'True
            Caption         =   "データービット："
            Height          =   360
            Index           =   2
            Left            =   300
            TabIndex        =   7
            Top             =   540
            Width           =   2100
         End
         Begin VB.Label lbRs 
            AutoSize        =   -1  'True
            Caption         =   "パリティ："
            Height          =   360
            Index           =   3
            Left            =   300
            TabIndex        =   9
            Top             =   1020
            Width           =   2100
         End
         Begin VB.Label lbRs 
            AutoSize        =   -1  'True
            Caption         =   "ストップビット："
            Height          =   360
            Index           =   4
            Left            =   300
            TabIndex        =   11
            Top             =   1500
            Width           =   2100
         End
      End
      Begin VB.Label lbRs 
         AutoSize        =   -1  'True
         Caption         =   "通信速度："
         Height          =   360
         Index           =   1
         Left            =   420
         TabIndex        =   4
         Top             =   1140
         Width           =   2160
      End
      Begin VB.Label lbRs 
         AutoSize        =   -1  'True
         Caption         =   "ポート番号："
         Height          =   360
         Index           =   0
         Left            =   420
         TabIndex        =   2
         Top             =   660
         Width           =   2160
      End
   End
   Begin VB.CommandButton btReturn 
      Cancel          =   -1  'True
      Caption         =   $"frmQRSettings.frx":01F7
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5100
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3300
      Width           =   1950
   End
End
Attribute VB_Name = "frmQRSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' プロパティ定義用
'----------------------------------------
Private fstAct As Boolean   'フォーム初回Active

'================================================================================
' 概要　：フォーム　Loadイベント
'================================================================================
Private Sub Form_Load()

  'フラグを初期化します
  IsFirstActivate = True

End Sub

'================================================================================
' 概要　：フォーム　Activateイベント
'================================================================================
Private Sub Form_Activate()

  'アクティベート初回時の場合
  If (IsFirstActivate) Then
    'コントロールの使用可否状態を設定します
    Call SetControlEnabled

    'フラグを更新します
    IsFirstActivate = False
  End If

End Sub

'================================================================================
' 概要　：フォーム　KeyPressイベント
'================================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)

  Select Case KeyAscii
    '[Enter]キーでフォーカスを移動
    Case vbKeyReturn
      Call SendKeys("{TAB}")
      KeyAscii = 0
      Exit Sub
  End Select

End Sub

'================================================================================
' 概要　：「設定・戻る」ボタン　Clickイベント
'================================================================================
Private Sub btReturn_Click()

  '画面を隠します
  'アンロードは呼び出し元で行います
  Me.Hide

End Sub

'================================================================================
' 概要　：「QRコードリーダーを使う」チェックボックス　Clickイベント
'================================================================================
Private Sub chkQREnabled_Click()

  'コントロールの使用可否状態を設定します
  Call SetControlEnabled

End Sub

'================================================================================
' 概要　：コントロールの使用可否状態を設定します
' 引数　：なし
' 戻り値：なし
'================================================================================
Private Sub SetControlEnabled()

  '「QRコードリーダーを使う」にチェックが入っている場合
  If (chkQREnabled.Value) Then
    'その他のコントロールを活性状態にします
    fraEnabled.Enabled = True

  '「QRコードリーダーを使う」にチェックが入っていない場合
  Else
    'その他のコントロールを非活性状態にします
    fraEnabled.Enabled = False
  End If

End Sub

'================================================================================
' 概要　：QRコードリーダー使用可否　プロパティ
'================================================================================
Public Property Get QREnabled() As Boolean

  '「QRコードリーダーを使う」にチェックが入っていればTrue、入っていなければFalseを返します
  If (chkQREnabled.Value = vbUnchecked) Then
    QREnabled = False
  Else
    QREnabled = True
  End If

End Property

Public Property Let QREnabled(ByVal Value As Boolean)

  If (Value) Then
    chkQREnabled.Value = vbChecked
  Else
    chkQREnabled.Value = vbUnchecked
  End If

End Property

'================================================================================
' 概要　：ポート番号　プロパティ
'================================================================================
Public Property Get portNo() As Integer

  'フォーム上で選択されている値をそのまま返します
  portNo = CInt(Replace(cboPortNo.Text, "Com", ""))

End Property

Public Property Let portNo(ByVal Value As Integer)

  Dim selIdx As Integer
  selIdx = 0  '初期値は"COM1"

  Dim i As Integer
  For i = 0 To (cboPortNo.ListCount - 1)
    If (Replace(cboPortNo.List(i), "Com", "") = Value) Then
      selIdx = i
      Exit For
    End If
  Next i

  cboPortNo.ListIndex = selIdx

End Property

'================================================================================
' 概要　：通信速度　プロパティ
'================================================================================
Public Property Get speed() As Long

  'フォーム上で選択されている値をそのまま返します
  speed = CLng(cboSpeed.Text)

End Property

Public Property Let speed(ByVal Value As Long)

  Dim selIdx As Integer
  selIdx = 6  '初期値は"9600"

  Dim i As Integer
  For i = 0 To (cboSpeed.ListCount - 1)
    If (cboSpeed.List(i) = Value) Then
      selIdx = i
      Exit For
    End If
  Next i

  cboSpeed.ListIndex = selIdx

End Property

'================================================================================
' 概要　：データビット　プロパティ
'================================================================================
Public Property Get dataBits() As Integer

  'フォーム上で選択されている値をそのまま返します
  dataBits = CInt(cboDataBits.Text)

End Property

Public Property Let dataBits(ByVal Value As Integer)

  Dim selIdx As Integer
  selIdx = 4  '初期値は"8"

  Dim i As Integer
  For i = 0 To (cboDataBits.ListCount - 1)
    If (cboDataBits.List(i) = Value) Then
      selIdx = i
      Exit For
    End If
  Next i

  cboDataBits.ListIndex = selIdx

End Property

'================================================================================
' 概要　：パリティ　プロパティ
'================================================================================
Public Property Get parity() As String

  'フォーム上で選択されている値の先頭1文字を返します
  parity = Left(cboParity.Text, 1)

End Property

Public Property Let parity(ByVal Value As String)

  Dim selIdx As Integer
  selIdx = 2  '初期値は"None"

  Dim i As Integer
  For i = 0 To (cboParity.ListCount - 1)
    If (Left(cboParity.List(i), 1) = Value) Then
      selIdx = i
      Exit For
    End If
  Next i

  cboParity.ListIndex = selIdx

End Property

'================================================================================
' 概要　：ストップビット　プロパティ
'================================================================================
Public Property Get stopBits() As Single

  'フォーム上で選択されている値をそのまま返します
  stopBits = CSng(cboStopBits.Text)

End Property

Public Property Let stopBits(ByVal Value As Single)

  Dim selIdx As Integer
  selIdx = 0  '初期値は"1"

  Dim i As Integer
  For i = 0 To (cboStopBits.ListCount - 1)
    If (cboStopBits.List(i) = Value) Then
      selIdx = i
      Exit For
    End If
  Next i

  cboStopBits.ListIndex = selIdx

End Property

'================================================================================
' 概要　：初回アクティベートフラグ　プロパティ
'================================================================================
Private Property Get IsFirstActivate() As Boolean
  IsFirstActivate = fstAct
End Property

Private Property Let IsFirstActivate(ByVal Value As Boolean)
  fstAct = Value
End Property
