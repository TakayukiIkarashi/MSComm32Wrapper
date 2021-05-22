Attribute VB_Name = "basMSComm32"
Option Explicit

'----------------------------------------
' �\���̒�`
'----------------------------------------
'QR�R�[�h���[�_�[�̏�Ԃ��i�[����\����
Public Type ComTable
  QREnabled As Boolean                  'QR�R�[�h�g�p��
  portNo As Integer                     '�|�[�g�ԍ�
  speed As Long                         '�ʐM���x
  parity As String                      '�p���e�B
  dataBits As Integer                   '�f�[�^�r�b�g
  stopBits As Single                    '�X�g�b�v�r�b�g
End Type
Public ComInfo As ComTable

'�v���W�F�N�g���ŗp�ӂ����G���[
Public Enum ErrNumber
  CannotPortOpen = 513                  '�|�[�g���J���Ȃ�
  UnknownEvent = 514                    '�F���ł��Ȃ��C�x���g�G���[
End Enum

'----------------------------------------
' INI�t�@�C����`
'----------------------------------------
Public Const INIFILE_NAME_ME As String = "mscomm32wrapper_rs232c.ini" 'INI�t�@�C����
Public Const INISEC_RS232C As String = "RS232C"                       '�Z�N�V�����FRS232C
Public Const INIKEY_RS232C_ENABLED As String = "ENABLED"              '�@QR�R�[�h���[�_�[�g�p��
Public Const INIKEY_RS232C_PORTNO As String = "PORTNO"                '�@�|�[�g�ԍ�
Public Const INIKEY_RS232C_SPEED As String = "SPEED"                  '�@�ʐM���x
Public Const INIKEY_RS232C_DATABITS As String = "DATABITS"            '�@�f�[�^�r�b�g
Public Const INIKEY_RS232C_PARITY As String = "PARITY"                '�@�p���e�B
Public Const INIKEY_RS232C_STOPBITS As String = "STOPBITS"            '�@�X�g�b�v�r�b�g

'================================================================================
' �T�v�@�F���̃A�v���P�[�V�����Ŏg�p����INI�t�@�C����ǂݍ��݁A
' �@�@�@�@�O��̋N����Ԃ𕜌����܂�
' �����@�F�Ȃ�
' �߂�l�F�Ȃ�
'================================================================================
Public Sub GetIniFile()

  '[QR�R�[�h���[�_�[���g��]
  Dim qrEn As Boolean
  qrEn = CBool(getINIValue(INIKEY_RS232C_ENABLED, INISEC_RS232C, "0", App.Path & "\" & INIFILE_NAME_ME))

  '[�|�[�g�ԍ�]
  Dim portNo As Integer
  portNo = Val(getINIValue(INIKEY_RS232C_PORTNO, INISEC_RS232C, "", App.Path & "\" & INIFILE_NAME_ME))
  If (portNo = 0) Then
    portNo = 1
  End If

  '[�ʐM���x]
  Dim speed As Long
  speed = Val(getINIValue(INIKEY_RS232C_SPEED, INISEC_RS232C, "", App.Path & "\" & INIFILE_NAME_ME))
  If (speed = 0) Then
    speed = 9600
  End If

  '[�p���e�B]
  Dim parity As String
  parity = getINIValue(INIKEY_RS232C_PARITY, INISEC_RS232C, "", App.Path & "\" & INIFILE_NAME_ME)
  If (parity = "") Then
    parity = "N"
  End If

  '[�f�[�^�r�b�g]
  Dim dataBits As Integer
  dataBits = Val(getINIValue(INIKEY_RS232C_DATABITS, INISEC_RS232C, "8", App.Path & "\" & INIFILE_NAME_ME))

  '[�X�g�b�v�r�b�g]
  Dim stopBits As Single
  stopBits = Val(getINIValue(INIKEY_RS232C_STOPBITS, INISEC_RS232C, "1", App.Path & "\" & INIFILE_NAME_ME))

  '�\���̂�RS232C�����Z�b�g���܂�
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
' �T�v�@�F���̃A�v���P�[�V�����Ŏg�p����INI�t�@�C����RS232C�ݒ�����������݂܂�
' �����@�F�Ȃ�
' �߂�l�F�Ȃ�
'================================================================================
Public Sub SetIniFile()

  '[QR�R�[�h���[�_�[���g��]
  Call setINIValue(ComInfo.QREnabled, INIKEY_RS232C_ENABLED, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[�|�[�g�ԍ�]
  Call setINIValue(ComInfo.portNo, INIKEY_RS232C_PORTNO, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[�ʐM���x]
  Call setINIValue(ComInfo.speed, INIKEY_RS232C_SPEED, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[�f�[�^�r�b�g]
  Call setINIValue(ComInfo.dataBits, INIKEY_RS232C_DATABITS, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[�p���e�B]
  Call setINIValue(ComInfo.parity, INIKEY_RS232C_PARITY, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

  '[�X�g�b�v�r�b�g]
  Call setINIValue(ComInfo.stopBits, INIKEY_RS232C_STOPBITS, INISEC_RS232C, App.Path & "\" & INIFILE_NAME_ME)

End Sub
