Attribute VB_Name = "Mod03_COMMON"
Option Explicit

'OS���ʎq
Public Enum OS_TYPE
    T01_MAC = 1
    T02_WIN = 2
    T03_OTHER = 0
End Enum

'�����R�[�h
Public Const P_vbLf = 10 '���C���t�B�[�h����
Public Const P_vbCr = 13 '�L�����b�W���^�[������
Public Const P_vbTab = 8 '�^�u����

'=================================
'Excel���s����OS�𒲂ׂĕԂ�
'����(1)�F�擾����OS�o�[�W�����i�ȗ��j
'�߂�l�FOS���ʎq�iEnum�Œ�`�������́j
'=================================
Public Function Get_MyOS(Optional ByRef str_ResOS As String) As Integer
    Dim str_OS As String 'OS
    Dim int_myOS As Integer

    'OS�̃o�[�W�������擾����
    str_OS = Application.OperatingSystem
    str_ResOS = str_OS

    'OS���ʎq���Z�b�g
    Select Case True
     Case str_OS Like "*Mac*"
        int_myOS = OS_TYPE.T01_MAC
     Case str_OS Like "*Windows*"
        int_myOS = OS_TYPE.T02_WIN
    Case Else
        int_myOS = OS_TYPE.T03_OTHER
    End Select

    Get_MyOS = int_myOS
End Function
 
'=================================
'���s����Excel�o�[�W�����𒲂ׂĕԂ�
'����(1)�F�Ȃ�
'�߂�l�FOS���ʎq�iEnum�Œ�`�������́j
'=================================
Public Function GET_MyExcelVersion() As String
    Dim str_VER As String 'OS

    'OS�̃o�[�W�������擾����
    str_VER = Application.Version

    GET_MyExcelVersion = str_VER
End Function
 
 
'=================================
'���s����OS�ʂ̉��s�R�[�h�𒲂ׂĕԂ�
'����(1)�F�Ȃ�
'�߂�l�F���s����
'=================================
Public Function GET_LineBreak() As String
    Dim str_LineBreak As String '���s

    Select Case Get_MyOS
        Case OS_TYPE.T01_MAC 'Mac OS
            str_LineBreak = Chr(P_vbCr) 'vbCr
        Case OS_TYPE.T02_WIN 'Windows
            str_LineBreak = Chr(P_vbCr) + Chr(P_vbLf) 'vbCrLf
        Case OS_TYPE.T03_OTHER
            str_LineBreak = Chr(P_vbLf)
    End Select
        
    GET_LineBreak = str_LineBreak
End Function

'==============================
'�A�N�e�B�u�Z������ʂ̍���[�ɕ\������
'����(1)�F�ΏۃV�[�g
'����(2)�F�ΏۃZ���i���[�ɕ\���������Z���j
'����(3)�F�G���[���b�Z�[�W�i�ȗ��j
'�߂�l�F�Ȃ�
'==============================
Public Sub Set_ScrollSheet(ByRef wkSheet As Worksheet _
                                , ByVal tagetCell As String _
                                , Optional ByRef errMsg As String)
    On Error GoTo EXIT_CODE
    '�V�[�g���A�N�e�B�u�ɂ��āA�w�肳�ꂽ�Z����I��
    With wkSheet
        .Activate
        .Range(tagetCell).Select
    End With
    '�A�N�e�B�u�Z���ւ̃X�N���[���ړ�
    With ActiveWindow
        .ScrollRow = ActiveCell.Row
        .ScrollColumn = ActiveCell.Column
    End With
EXIT_CODE:
    If Err.Description <> "" Then
        errMsg = errMsg & Err.Description & vbNewLine
    End If
End Sub
