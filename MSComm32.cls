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
' �ϐ���`
'----------------------------------------
Private MainForm As frmMSComm32         'MSComm32�t�H�[��

'================================================================================
' TextData�v���p�e�B
'================================================================================
Public Property Get TextData() As String
  TextData = MainForm.TextData
End Property

Public Property Let TextData(ByVal Value As String)
  MainForm.TextData = Value
End Property

'================================================================================
' PortOpen�v���p�e�B
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
' �R���X�g���N�^
'================================================================================
Private Sub Class_Initialize()
  '�ݒ�t�@�C������RS232C�ڑ��̐ݒ���擾���܂�
  Call GetIniFile

  '���C���t�H�[�������[�h���܂�
  Set MainForm = New frmMSComm32
  MainForm.Visible = False
End Sub

'================================================================================
' �f�X�g���N�^
'================================================================================
Private Sub Class_Terminate()
  'COM�|�[�g����܂�
  Call CloseComPort

  '���C���t�H�[������܂��B
  Unload MainForm
End Sub

'================================================================================
' �T�v�@�F�uRS232C�ݒ�v�t�H�[�����N�����܂�
' �����@�F�Ȃ�
' �߂�l�F�Ȃ�
'================================================================================
Public Sub ShowQRSettingsForm()

  '�uQR�ݒ�v�E�B���h�E���`���܂�
  Dim f As New frmQRSettings
  f.QREnabled = ComInfo.QREnabled
  f.portNo = ComInfo.portNo
  f.speed = ComInfo.speed
  f.dataBits = ComInfo.dataBits
  f.parity = ComInfo.parity
  f.stopBits = ComInfo.stopBits

  '�uQR�ݒ�v�E�B���h�E��\�����܂�
  f.Show vbModal

  '�uQR�ݒ�v�E�B���h�E�̓��e���\���̂ɔ��f���܂�
  ComInfo.QREnabled = f.QREnabled
  ComInfo.portNo = f.portNo
  ComInfo.speed = f.speed
  ComInfo.dataBits = f.dataBits
  ComInfo.parity = f.parity
  ComInfo.stopBits = f.stopBits

  '�uQR�ݒ�v�E�B���h�E����܂�
  Unload f

  'RS232C�̐ݒ��ݒ�t�@�C���ɕۑ����܂��B
  Call SetIniFile

End Sub

'================================================================================
' �T�v�@�FCOM�|�[�g���J���܂�
' �����@�F�Ȃ�
' �߂�l�F�Ȃ�
'================================================================================
Private Sub OpenComPort()

  '���ł�COM�|�[�g���J���Ă���ꍇ�͏����𔲂��܂�
  If (MainForm.MSComm1.PortOpen) Then
    Exit Sub
  End If

  On Error GoTo Exception

  With MainForm.MSComm1
    '�|�[�g�ԍ��̐ݒ�
    .CommPort = ComInfo.portNo

    'COM�|�[�g�ݒ蕶������i�[����ϐ����`���܂�
    Dim s As String
    s = ""
    s = s & CStr(ComInfo.speed) & ","
    s = s & ComInfo.parity & ","
    s = s & CStr(ComInfo.dataBits) & ","
    s = s & CStr(ComInfo.stopBits)

    'COM�|�[�g�ݒ蕶����
    .Settings = s

    'Input�v���p�e�B�g�p���Ƀo�b�t�@�S�̂�ǂݎ��悤�ɐݒ肵�܂�
    .InputLen = 0

    '�|�[�g���J���܂�
    .PortOpen = True
    .RTSEnable = True
  End With

  If (MainForm.MSComm1.PortOpen = False) Then
    Call Err.Raise(ErrNumber.CannotPortOpen, "OpenComPort", "�|�[�g���J���܂���B")
  End If

  Exit Sub

Exception:
  Call Err.Raise(Err.Number, "OpenComPort", Err.Description)

End Sub

'================================================================================
' �T�v�@�FCOM�|�[�g����܂�
' �����@�F�Ȃ�
' �߂�l�F�Ȃ�
'================================================================================
Private Sub CloseComPort()

  '���ł�COM�|�[�g�����Ă���Ώ����𔲂��܂�
  If (MainForm.MSComm1.PortOpen = False) Then
    Exit Sub
  End If

  On Error GoTo Exception

  '�|�[�g����܂�
  MainForm.MSComm1.RTSEnable = False
  MainForm.MSComm1.PortOpen = False

  Exit Sub

Exception:
  Call Err.Raise(Err.Number, "OpenComPort", Err.Description)

End Sub