VERSION 5.00
Begin VB.Form frmQRSettings 
   BorderStyle     =   4  '�Œ�°� ����޳
   Caption         =   "QR�ݒ�"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "���C���I"
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
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CheckBox chkQREnabled 
      Caption         =   "QR�R�[�h���[�_�[���g��"
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
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   5
         Top             =   1080
         Width           =   1620
      End
      Begin VB.ComboBox cboPortNo 
         Height          =   480
         ItemData        =   "frmQRSettings.frx":00DA
         Left            =   2700
         List            =   "frmQRSettings.frx":0115
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   3
         Top             =   600
         Width           =   1620
      End
      Begin VB.Frame faConnection 
         Caption         =   "�ڑ��ݒ�"
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
            Style           =   2  '��ۯ���޳� ؽ�
            TabIndex        =   8
            Top             =   480
            Width           =   1620
         End
         Begin VB.ComboBox cboParity 
            Height          =   480
            ItemData        =   "frmQRSettings.frx":01A6
            Left            =   2460
            List            =   "frmQRSettings.frx":01B9
            Style           =   2  '��ۯ���޳� ؽ�
            TabIndex        =   10
            Top             =   960
            Width           =   1620
         End
         Begin VB.ComboBox cboStopBits 
            Height          =   480
            ItemData        =   "frmQRSettings.frx":01DB
            Left            =   2460
            List            =   "frmQRSettings.frx":01E8
            Style           =   2  '��ۯ���޳� ؽ�
            TabIndex        =   12
            Top             =   1440
            Width           =   1620
         End
         Begin VB.Label lbRs 
            AutoSize        =   -1  'True
            Caption         =   "�f�[�^�[�r�b�g�F"
            Height          =   360
            Index           =   2
            Left            =   300
            TabIndex        =   7
            Top             =   540
            Width           =   2100
         End
         Begin VB.Label lbRs 
            AutoSize        =   -1  'True
            Caption         =   "�p���e�B�F"
            Height          =   360
            Index           =   3
            Left            =   300
            TabIndex        =   9
            Top             =   1020
            Width           =   2100
         End
         Begin VB.Label lbRs 
            AutoSize        =   -1  'True
            Caption         =   "�X�g�b�v�r�b�g�F"
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
         Caption         =   "�ʐM���x�F"
         Height          =   360
         Index           =   1
         Left            =   420
         TabIndex        =   4
         Top             =   1140
         Width           =   2160
      End
      Begin VB.Label lbRs 
         AutoSize        =   -1  'True
         Caption         =   "�|�[�g�ԍ��F"
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
         Name            =   "���C���I"
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
' �v���p�e�B��`�p
'----------------------------------------
Private fstAct As Boolean   '�t�H�[������Active

'================================================================================
' �T�v�@�F�t�H�[���@Load�C�x���g
'================================================================================
Private Sub Form_Load()

  '�t���O�����������܂�
  IsFirstActivate = True

End Sub

'================================================================================
' �T�v�@�F�t�H�[���@Activate�C�x���g
'================================================================================
Private Sub Form_Activate()

  '�A�N�e�B�x�[�g���񎞂̏ꍇ
  If (IsFirstActivate) Then
    '�R���g���[���̎g�p�ۏ�Ԃ�ݒ肵�܂�
    Call SetControlEnabled

    '�t���O���X�V���܂�
    IsFirstActivate = False
  End If

End Sub

'================================================================================
' �T�v�@�F�t�H�[���@KeyPress�C�x���g
'================================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)

  Select Case KeyAscii
    '[Enter]�L�[�Ńt�H�[�J�X���ړ�
    Case vbKeyReturn
      Call SendKeys("{TAB}")
      KeyAscii = 0
      Exit Sub
  End Select

End Sub

'================================================================================
' �T�v�@�F�u�ݒ�E�߂�v�{�^���@Click�C�x���g
'================================================================================
Private Sub btReturn_Click()

  '��ʂ��B���܂�
  '�A�����[�h�͌Ăяo�����ōs���܂�
  Me.Hide

End Sub

'================================================================================
' �T�v�@�F�uQR�R�[�h���[�_�[���g���v�`�F�b�N�{�b�N�X�@Click�C�x���g
'================================================================================
Private Sub chkQREnabled_Click()

  '�R���g���[���̎g�p�ۏ�Ԃ�ݒ肵�܂�
  Call SetControlEnabled

End Sub

'================================================================================
' �T�v�@�F�R���g���[���̎g�p�ۏ�Ԃ�ݒ肵�܂�
' �����@�F�Ȃ�
' �߂�l�F�Ȃ�
'================================================================================
Private Sub SetControlEnabled()

  '�uQR�R�[�h���[�_�[���g���v�Ƀ`�F�b�N�������Ă���ꍇ
  If (chkQREnabled.Value) Then
    '���̑��̃R���g���[����������Ԃɂ��܂�
    fraEnabled.Enabled = True

  '�uQR�R�[�h���[�_�[���g���v�Ƀ`�F�b�N�������Ă��Ȃ��ꍇ
  Else
    '���̑��̃R���g���[����񊈐���Ԃɂ��܂�
    fraEnabled.Enabled = False
  End If

End Sub

'================================================================================
' �T�v�@�FQR�R�[�h���[�_�[�g�p�ہ@�v���p�e�B
'================================================================================
Public Property Get QREnabled() As Boolean

  '�uQR�R�[�h���[�_�[���g���v�Ƀ`�F�b�N�������Ă����True�A�����Ă��Ȃ����False��Ԃ��܂�
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
' �T�v�@�F�|�[�g�ԍ��@�v���p�e�B
'================================================================================
Public Property Get portNo() As Integer

  '�t�H�[����őI������Ă���l�����̂܂ܕԂ��܂�
  portNo = CInt(Replace(cboPortNo.Text, "Com", ""))

End Property

Public Property Let portNo(ByVal Value As Integer)

  Dim selIdx As Integer
  selIdx = 0  '�����l��"COM1"

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
' �T�v�@�F�ʐM���x�@�v���p�e�B
'================================================================================
Public Property Get speed() As Long

  '�t�H�[����őI������Ă���l�����̂܂ܕԂ��܂�
  speed = CLng(cboSpeed.Text)

End Property

Public Property Let speed(ByVal Value As Long)

  Dim selIdx As Integer
  selIdx = 6  '�����l��"9600"

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
' �T�v�@�F�f�[�^�r�b�g�@�v���p�e�B
'================================================================================
Public Property Get dataBits() As Integer

  '�t�H�[����őI������Ă���l�����̂܂ܕԂ��܂�
  dataBits = CInt(cboDataBits.Text)

End Property

Public Property Let dataBits(ByVal Value As Integer)

  Dim selIdx As Integer
  selIdx = 4  '�����l��"8"

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
' �T�v�@�F�p���e�B�@�v���p�e�B
'================================================================================
Public Property Get parity() As String

  '�t�H�[����őI������Ă���l�̐擪1������Ԃ��܂�
  parity = Left(cboParity.Text, 1)

End Property

Public Property Let parity(ByVal Value As String)

  Dim selIdx As Integer
  selIdx = 2  '�����l��"None"

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
' �T�v�@�F�X�g�b�v�r�b�g�@�v���p�e�B
'================================================================================
Public Property Get stopBits() As Single

  '�t�H�[����őI������Ă���l�����̂܂ܕԂ��܂�
  stopBits = CSng(cboStopBits.Text)

End Property

Public Property Let stopBits(ByVal Value As Single)

  Dim selIdx As Integer
  selIdx = 0  '�����l��"1"

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
' �T�v�@�F����A�N�e�B�x�[�g�t���O�@�v���p�e�B
'================================================================================
Private Property Get IsFirstActivate() As Boolean
  IsFirstActivate = fstAct
End Property

Private Property Let IsFirstActivate(ByVal Value As Boolean)
  fstAct = Value
End Property
