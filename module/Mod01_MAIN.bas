Attribute VB_Name = "Mod01_Main"

'==========================================
' �yCREATE SQL�z BUTTON CLICK
'==========================================
Public Sub CreateSQL_ButtonClick()
    Dim isOK As Boolean                  '���s����
    Dim errMsg As String                 '�G���[���b�Z�[�W
    Dim mySheet As Worksheet        '�ΏۃV�[�g
    Dim lng_stRow As Long               '�J�������̓��͊J�n�s
    Dim lng_edRow    As Long         '�J�������̓��͍ŉ��s ���A�������Z���͈͓��i��������j
    Dim str_SQL As String                   'SQL
    Dim myTableInfo As TABLE_INFO  '���͂��ꂽ�e�[�u�����
    Dim resSet As RESULT_SET 'SQL�o�͌��ʃZ�b�g
    Dim str_msgText As String '���s���ʁi�ڍׁj
    
    On Error GoTo EXIT_CODE
    
    Set mySheet = ActiveSheet '�A�N�e�B�u�V�[�g����ƃV�[�g�ɃZ�b�g
    lng_stRow = P_DATA_HEAD_ROW + 1  '�J�������̓��͊J�n�s�i�����o���s�{1�j
    
    '���s���ʂ̃N���A
    Call Clear_ResultCells(mySheet, P_SQL_START_ROW, CLng(Rows.Count))

    '�V�[�g�ɓ��͂��ꂽ�e�[�u�����̎擾
    isOK = Get_TableInfo(mySheet, myTableInfo)
    If isOK = False Then
        GoTo EXIT_CODE
    End If
    
    '�Z���̓��͒l�`�F�b�N�̎��s
    isOK = IsOK_ValidationCheck(mySheet, lng_stRow, lng_edRow, myTableInfo, errMsg)
    If isOK = False Then
        GoTo EXIT_CODE
    End If
    
    ' CREATE TABLE SQL���̍쐬
    str_SQL = Return_SQL_CREATE_TABLE(mySheet, lng_stRow, lng_edRow, myTableInfo, errMsg)
    If errMsg <> "" Then
         isOK = False
    Else
        isOK = True
    End If
    
    'SQL�����V�[�g�ɏo��
    isOK = IsOK_WriteSQL(mySheet, str_SQL, resSet, errMsg)
    If isOK = False Then
        GoTo EXIT_CODE
    End If

EXIT_CODE:
    ' ���b�Z�[�W�쐬 ----------------------
    Select Case isOK
        Case True
            str_msgText = "SQL���̍쐬�ɐ������܂����I"
        Case False
            str_msgText = "SQL���̍쐬�Ɏ��s���܂����I" & vbNewLine & errMsg
    End Select
    
    '���s���ʂ̎擾
    With resSet
        '��������
        .dat_resTime = Now()
        '���s����
        .str_resultOK = IIf(isOK = True, "OK", "NG")
        '���s�ڍ�
        .str_detailMsg = str_msgText
    End With
    
    '���s���ʂ̏�������
    With mySheet
        .Range(P_RESULT_DATE_CELL).Value = Format(resSet.dat_resTime, "YYYY/M/D HH:MM")
        .Range(P_RESULT_OKNG_CELL).Value = resSet.str_resultOK
        .Range(P_RESULT_DETAIL_CELL).Value = resSet.str_detailMsg
    End With
    
    ' ���b�Z�[�W�o�� ----------------------
    Select Case isOK
        Case True
        'OK
            Call Set_ScrollSheet(mySheet, "R1") 'SQL�쐬���ʂփX�N���[�����ړ�
            MsgBox str_msgText, vbInformation, "���s���� [" & mySheet.Name & "]�V�[�g "
        Case False
        'NG
            Call Set_ScrollSheet(ActiveSheet, "A1") '�e�[�u����`�����͂փX�N���[�����ړ�
            MsgBox str_msgText, vbExclamation, "���s���ʁi�G���[�j[" & mySheet.Name & "]�V�[�g "
    End Select
    
    Set mySheet = Nothing
End Sub

'==========================================
' �yBack InputForm�z BUTTON CLICK
'==========================================
Public Sub BackInputForm_ButtonClick()
    '�V�[�g�ړ�
    Call Set_ScrollSheet(ActiveSheet, "A1") '�e�[�u����`�����͂փX�N���[�����ړ�
End Sub

'==========================================
' �yCOPY CREATESQL�z BUTTON CLICK
'==========================================
Public Sub CopyCreateSQL_ButtonClick()
    Dim isOK As Boolean                  '���s����
    Dim errMsg As String                 '�G���[���b�Z�[�W
    Dim mySheet As Worksheet        '�ΏۃV�[�g
    Dim lng_stRow As Long               'SQL�̏o�͊J�n�s
    Dim lng_edRow    As Long         'SQL�̏o�͍ŉ��s ���A�������Z���͈͓��i��������j

    On Error GoTo EXIT_CODE
    
    Set mySheet = ActiveSheet '�A�N�e�B�u�V�[�g����ƃV�[�g�ɃZ�b�g
    lng_stRow = P_SQL_START_ROW  'SQL�̏o�͊J�n�s
    With mySheet
        lng_edRow = .Cells(Rows.Count, .Range(P_SQL_COLUMN & "1").Column).End(xlUp).Row             '�A�������Ώۗ�̍ŉ��s
        If lng_stRow < lng_edRow Then
            .Range(P_SQL_COLUMN & lng_stRow & ":" & P_SQL_COLUMN & lng_edRow).Copy
            isOK = True
        Else
EXIT_CODE:
            errMsg = Err.Description
            isOK = False
        End If
    End With
    
    Select Case isOK
        Case True
            MsgBox "OK", vbInformation, "RESULT"
        Case False
            MsgBox "NG" & vbNewLine & errMsg, vbExclamation, "RESULT"
    End Select

    Set mySheet = Nothing
End Sub
