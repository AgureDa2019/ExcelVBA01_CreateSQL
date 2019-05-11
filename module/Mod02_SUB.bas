Attribute VB_Name = "Mod02_SUB"
Option Explicit

'==========================================
'�e�[�u�����̎擾
'����(1)�F�@�ΏۃV�[�g (IN)
'����(2)�F�@�e�[�u�����̎擾�f�[�^ (OUT)
'����(3)�F�@�G���[���b�Z�[�W (OUT)
'�߂�l�@���ʁiOK�FTrue�ANG�FFalse�j
'==========================================
Public Function Get_TableInfo(ByRef wkSheet As Worksheet _
                                         , ByRef myTableInfo As TABLE_INFO _
                                         , Optional errMsg As String) As Boolean
Dim isOK As Boolean

On Error GoTo ERR_CODE
    
    With myTableInfo
        ' D��
        .str_DbName = wkSheet.Range(P_DBNAME_CELL).Value
        .str_TblName = wkSheet.Range(P_TABLENAME_CELL).Value
        .str_TblComment = wkSheet.Range(P_TABLECOMMENT_CELL).Value
        ' F��
        .str_TblType = wkSheet.Range(P_TABLETYPE_CELL).Value
        .str_MySql_ver = wkSheet.Range(P_MYSQLVER_CELL).Value
        .str_ENGINE = wkSheet.Range(P_ENGINE_CELL).Value
        .str_CHARSET = wkSheet.Range(P_CHARSET_CELL).Value
        '�ꎞ�e�[�u�����ǂ���
        If .str_TblType Like "*TEMPORARY*" Then
            .is_TEMPORARY_Table = True
        Else
            .is_TEMPORARY_Table = False
        End If
        .str_PURPOSE = wkSheet.Range(P_PURPOSE_CELL).Value
    End With
    
ERR_CODE:
    If Err.Description <> "" Then
        errMsg = errMsg & " [ERROR] GET_TABLE_INFO " & vbNewLine
        errMsg = errMsg & "[ERROR NUMBER] " & Err.Number & vbNewLine
        errMsg = errMsg & "[DETAIL] " & Err.Description
        isOK = False
    Else
     isOK = True
    End If
    
   Get_TableInfo = isOK
End Function


'==========================================
'���s���ʂ̃N���A
'����(1)�F�@�ΏۃV�[�g
'����(2)�F�@���s���ʂ̓��͊J�n�s
'����(3)�F�@���s���ʂ̃V�[�g�ŉ��s
'�߂�l�@�Ȃ�
'==========================================
Public Sub Clear_ResultCells(ByRef mySheet As Worksheet _
                                        , ByRef lng_stRow As Long _
                                        , ByRef lng_edRow As Long)
    '--------------------------
    '�V�[�g���̃Z���l�N���A
    '--------------------------
    With mySheet
    
    
        'No
        .Range(P_NO_COLUMN & P_DATA_START_ROW & ":" & P_NO_COLUMN & lng_edRow).Value = ""
        '�o�͌���
        .Range(P_RESULT_RENGE).Value = "" '��������
        .Range(P_SQL_COLUMN & lng_stRow & ":" & P_SQL_COLUMN & lng_edRow).Value = "" 'SQL��
        
        '-------------------------------
        '�w�i�F�������N���A: �e�[�u�����
        '-------------------------------
        '�X�L�[�}�� *
        '�����e�[�u���� *
        '�_���e�[�u���� *
        .Range(P_DBNAME_CELL _
                   & ":" & P_TABLECOMMENT_CELL).Interior.ColorIndex = 0 '�w�i�F�������N���A
        '���p�ړI *
        .Range(P_PURPOSE_CELL).Interior.ColorIndex = 0 '�w�i�F�������N���A
        '�e�[�u���敪 *
        'MySQL ver *
        'ENGINE *
        'CHARSET *
        .Range(P_TABLETYPE_CELL _
                   & ":" & P_CHARSET_CELL).Interior.ColorIndex = 0 '�w�i�F�������N���A
        
        '-------------------------------
        '�w�i�F�������N���A: �J�������
        '-------------------------------
        .Range(P_PRIMARY_COLUMN & P_DATA_START_ROW _
                   & ":" & P_INDEX3_COLUMN & P_DATA_MAX_ROW).Interior.ColorIndex = 0 '�w�i�F�������N���A
    End With
    
End Sub

'==========================================
'�e�[�u����`�@���͒l�`�F�b�N    ## (unfinished) ##
'����(1)�F�@�ΏۃV�[�g
'����(2)�F�@�J�������̓��͊J�n�s
'����(3)�F�@�J�������̓��͍ŉ��s
'����(4)�F�@�e�[�u�����(IN)
'����(5)�F�@�G���[���b�Z�[�W(OUT)
'�߂�l�@�`�F�b�N���ʁiOK�FTrue�ANG�FFalse�j
'==========================================
Public Function IsOK_ValidationCheck(ByRef mySheet As Worksheet _
                                                     , ByRef lng_stRow As Long _
                                                     , ByRef lng_edRow As Long _
                                                     , ByRef myTableInfo As TABLE_INFO _
                                                     , ByRef errMsg As String) As Boolean
    Dim isOK                 As Boolean   '�`�F�b�N����
    Dim lng_TotalCnt     As Long         '�w�肵���Z���͈͓��̓��͌���
    Dim lng_OutOfRengeCnt As Long   '�͈͊O�̓��͌���
    '�`�F�b�N�s
    Dim lng_nowRow As Long '�J�������̓��̓`�F�b�N�s
    Dim str_targetCell As String  '�`�F�b�N�ΏۃZ��
    '�Z���͈�
    Dim str_colNameRange As String '�J�������̓��̓Z���͈�
    Dim str_OutOfRenge As String     '�J�������̓��̓Z���͈͊O

    errMsg = ""
    
    On Error GoTo EXIT_CODE
    
    '-----------------------------------------------
    '(1/5) �V�[�g�̌`���`�F�b�N ���^�C�g�������������`�F�b�N
    '-----------------------------------------------
    If mySheet.Range(P_LABEL_TABLE_CELL).Value <> "�e�[�u�����" _
     Or mySheet.Range(P_LABEL_COLUMN_CELL).Value <> "�J�������" _
     Or mySheet.Range(P_NO_COLUMN & P_DATA_HEAD_ROW).Value <> "No" Then
        errMsg = errMsg & "�V�[�g�̃t�H�[�}�b�g������������܂���B�����`�V�[�g���R�s�[���čĎ��s���Ă��������B"
        GoTo EXIT_CODE
    End If

    '-----------------------------------------------
    '(2/5) �J�������̓��͌����`�F�b�N
    '-----------------------------------------------
    str_colNameRange = P_COLNAME_COLUMN & P_DATA_START_ROW & ":" & P_COLNAME_COLUMN & P_DATA_MAX_ROW
    With mySheet
        lng_TotalCnt = WorksheetFunction.CountA(.Range(str_colNameRange))     '�Z���͈͓��̓��͌���
        lng_edRow = .Cells(Rows.Count, .Range(P_COLNAME_COLUMN & "1").Column).End(xlUp).Row          '�A������3��ځi��������j�̍ŉ��s
    End With
    '(2/5) - 1
    If lng_TotalCnt < 1 Then
        errMsg = errMsg & "�J��������[Column Name]�񂪖����͂ł��B" & "�`�F�b�N�͈́F" & str_colNameRange & vbNewLine
    End If
    '(2/5) - 2
    If (lng_TotalCnt + P_DATA_HEAD_ROW) <> lng_edRow Then
        errMsg = errMsg & "�J��������[Column Name]�͘A�����ē��͂��Ă��������B" & "�`�F�b�N�͈́F" & str_colNameRange & vbNewLine
    End If
    '(2/5) - 3
    If lng_edRow < P_DATA_MAX_ROW Then
        str_OutOfRenge = P_COLNAME_COLUMN & (lng_edRow + 1) & ":" & P_INDEX3_COLUMN & P_DATA_MAX_ROW
        lng_OutOfRengeCnt = WorksheetFunction.CountA(mySheet.Range(str_OutOfRenge))  '�Z���͈͊O�̓��͌���
        If lng_OutOfRengeCnt > 0 Then
            errMsg = errMsg & "�J�������ɕs�v�ȃf�[�^���܂܂�Ă��܂��B" & "�`�F�b�N�͈́F" & str_OutOfRenge & vbNewLine
        End If
    End If
    '(2/5) �G���[�̂��ߍ�Ƃ𒆎~
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    
    '-------------------------------------
    '(3/5) �e�[�u�����̃`�F�b�N
    '--------------------------------------
    '---------------------------------
    ' (3/5)-1 �e�[�u����� �K�{����
    '---------------------------------
    '�K�{(1):[�e�[�u���敪] *
    Call IsOK_InputRange(mySheet, P_TABLETYPE_CELL, "�e�[�u����� ��[�e�[�u���敪]", errMsg)
    '�K�{(2):[�X�L�[�}��] *
    Call IsOK_InputRange(mySheet, P_DBNAME_CELL, "�e�[�u����� ��[�X�L�[�}��]", errMsg)
    '�K�{(3):[�����e�[�u����] *
    Call IsOK_InputRange(mySheet, P_TABLENAME_CELL, "�e�[�u����� ��[�����e�[�u����]", errMsg)
    '�K�{(4):[�_���e�[�u����] *
    Call IsOK_InputRange(mySheet, P_TABLECOMMENT_CELL, "�e�[�u����� ��[�_���e�[�u����]", errMsg)
    '�K�{(5):[MySQL ver] *
    Call IsOK_InputRange(mySheet, P_MYSQLVER_CELL, "�e�[�u����� ��[MySQL ver]", errMsg)
    '�K�{(6):[ENGINE] *
    Call IsOK_InputRange(mySheet, P_ENGINE_CELL, "�e�[�u����� ��[ENGINE]", errMsg)
     '�K�{(7):[CHARSET] *
    Call IsOK_InputRange(mySheet, P_CHARSET_CELL, "�e�[�u����� ��[CHARSET]", errMsg)
     '�K�{(8):[���p�ړI] *
    Call IsOK_InputRange(mySheet, P_PURPOSE_CELL, "�e�[�u����� ��[���p�ړI]", errMsg)
    '(3/5) �G���[�̂��ߍ�Ƃ𒆎~
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    
    '-----------------------------------------------
    '(4/5)  �J�������̃`�F�b�N
    '-----------------------------------------------
    For lng_nowRow = lng_stRow To lng_edRow
        With mySheet
            '-----------------------------------------------
            ' (4/5)-1 �J������� �K�{����
            '-----------------------------------------------
            '�K�{(1):[�����J������] *
            str_targetCell = P_COLNAME_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��[�����J������]", errMsg)
            '�K�{(1):[�_���J������] *
            str_targetCell = P_COLCOMMENT_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��[�_���J������]", errMsg)
            '�K�{(1):[�^����] *
            str_targetCell = P_TYPE_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��[�^����]", errMsg)
            '�K�{(1):[�f�[�^�^] *
            str_targetCell = P_TYPEDETAIL_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��[�f�[�^�^]", errMsg)
            
            '-----------------------------------------------
            '(4/5)-2 �^�`�F�b�N
            '-----------------------------------------------
            ' --- ���l�^ / ������^---
            Select Case True
                Case CStr(.Range(P_TYPEDETAIL_COLUMN & lng_nowRow).Value) Like "*(M)*"
                '�f�[�^�^��(M)���܂ޏꍇ�A(M)�͓��͕K�{
                    str_targetCell = P_M_COLUMN & lng_nowRow
                    Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��(M)", errMsg)
                Case CStr(.Range(P_TYPEDETAIL_COLUMN & lng_nowRow).Value) Like "*(M,D)*"
                '�f�[�^�^��(M,D)���܂ޏꍇ�A(M)(D)�͓��͕K�{
                    str_targetCell = P_M_COLUMN & lng_nowRow
                    Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��(M)", errMsg)
                    str_targetCell = P_D_COLUMN & lng_nowRow
                    Call IsOK_InputRange(mySheet, str_targetCell, "�J������� ��(D)", errMsg)
            End Select
            ' --- ���t�E���Ԍ^ ---
            ' (tsp)�͑ΏۊO
                  
            '-----------------------------------------------
            '(4/5)-3 �����`�F�b�N ## (unfinished) ##
            '-----------------------------------------------
            ' --- ���l�^ ---
            ' --- ���t�E���Ԍ^ ---
            ' --- ������^ ---
        End With
    Next lng_nowRow
    
    '(4/5) �G���[�̂��ߍ�Ƃ𒆎~
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    
    '-----------------------------------------------
    '(5/5) MYSQL Version 8.0 CHECK
    '-----------------------------------------------
    With myTableInfo
        If .str_MySql_ver Like "*8.0*" _
            And StrComp(.str_CHARSET, "utf8", vbTextCompare) = 0 Then
            ' �o�[�W������8�ŕ����R�[�h��utf8�̏ꍇ�Autf8mb4���I������Ă��Ȃ��ƃG���[����
            errMsg = errMsg & "MYSQL �̃o�[�W������8�ŕ����R�[�h��utf8�̏ꍇ�Autf8mb4��I�����Ă��������B" & "�`�F�b�N�͈́F" & P_PURPOSE_CELL & vbNewLine
        End If
    End With
    
    '(5/5) �G���[�̂��ߍ�Ƃ𒆎~
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    

     isOK = True
     
EXIT_CODE:
    IsOK_ValidationCheck = isOK
End Function

'====================================
'�K�{���ڂ̓��̓`�F�b�N�ƁA�G���[���͑ΏۃZ�����w��F�œh��Ԃ�
'����(1)�F�ΏۃV�[�g
'����(2)�F�ΏۃZ����
'����(3)�F�G���[���ɏo�͂��鍀�ږ�
'����(4)�F�G���[���b�Z�[�W
'�߂�l�F�`�F�b�N����
'====================================
Public Function IsOK_InputRange(ByRef mySheet As Worksheet _
                                                , ByRef str_CellName As String _
                                                , ByRef str_ItemName As String _
                                                , ByRef errMsg As String) As Boolean
    Dim isOK As Boolean
    Dim lng_targetColor As Long
    
    lng_targetColor = 44
    
    If Len(Trim(mySheet.Range(str_CellName).Value)) = 0 Then
    '�K�{���ڂ̃`�F�b�NNG
        errMsg = errMsg & str_ItemName & "�������͂ł��B" & "�`�F�b�N�͈́F" & str_CellName & "�Z��" & vbNewLine
        '�w��F�œh��Ԃ�
        Call Change_RangeBGColor(mySheet.Range(str_CellName), lng_targetColor)
        isOK = False
    Else
    '�`�F�b�NOK
        isOK = True
    End If
    
    IsOK_InputRange = isOK
End Function


'==========================================
'�ΏۃZ���͈͂̔w�i�F��h��Ԃ�
'����(1)�F �Ώ�RANGE
'����(2)�F �w�i�F�̃C���f�b�N�X�ԍ�   ex) RED=3
'�߂�l�@�Ȃ�
'==========================================
Public Sub Change_RangeBGColor(ByRef targetRange As Range _
                                            , ByRef lng_colorIndex As Long)
    targetRange.Interior.ColorIndex = lng_colorIndex
End Sub


'==========================================
'�e�[�u����` CREATE SQL���̍쐬
'����(1)�F�@�ΏۃV�[�g(IN/OUT)
'����(2)�F�@�J�������̓��͊J�n�s(IN)
'����(3)�F�@�J�������̓��͍ŉ��s(IN)
'����(4)�F�@�e�[�u�����(IN)
'����(5)�F�@�G���[���b�Z�[�W(OUT)
'�߂�l�@�`�F�b�N���ʁiOK�FTrue�ANG�FFalse�j
' * https://dev.mysql.com/doc/refman/5.6/ja/create-table.html
'==========================================
Public Function Return_SQL_CREATE_TABLE(ByRef wkSheet As Worksheet _
                                                               , ByVal lng_stRow As Long _
                                                               , ByVal lng_edRow As Long _
                                                               , ByRef myTableInfo As TABLE_INFO _
                                                               , ByRef errMsg As String) As String
    Dim str_SQL As String                 '�쐬����SQL��
    Dim lng_No As Long                  '���ݍs�̃J����No
    Dim lng_currentRow As Long      '���݂̍s�ԍ�
    Dim myCurrentSet As CURRENT_COLUMN_SET '���ݍs�̃J�������
    Dim str_CreateDefinition As String '�J������`��SQL
    Dim var_inputRange As Variant '���͔͈͓��f�[�^�i�z��Ƃ��Ĉꎞ�ێ��j
    Dim indexSet()  As COLUMN_LIST_SET 'INDEX1~3
    Dim primaryKeySet  As COLUMN_LIST_SET '��L�[PRIMARY Key
    Dim int_totalIndexCnt() As Integer '�C���f�b�N�X�̓��͌���
    Dim int_tmpIndexCnt As Integer '�C���f�b�N�X�̓��͌���
    Dim int_primaryKeyCnt As Integer   '��L�[�̓��͌���
    Dim int_loop As Integer
    Dim int_subLoop As Integer
    
    '������
    str_SQL = ""
    str_CreateDefinition = ""
    ReDim int_totalIndexCnt(1 To P_INDEX_COUNT)
    For int_loop = 1 To P_INDEX_COUNT
        int_totalIndexCnt(int_loop) = 0
    Next int_loop
    int_primaryKeyCnt = 0
    ReDim indexSet(1 To P_INDEX_COUNT)
    
    On Error GoTo EXIT_CODE
    
    '�J�������̓��͊J�n�s�ƍŉ��s���������͈͓����`�F�b�N
    If lng_stRow > P_DATA_HEAD_ROW _
       And lng_stRow < lng_edRow _
       And lng_edRow <= P_DATA_MAX_ROW Then
      'ROW CHECK OK
    Else
     'ROW CHECK NG
        errMsg = errMsg & "[ERROR] CREATE SQL ROW CHECK NG"
        GoTo EXIT_CODE
    End If
    
    'PRIMARY KEY �̓��͌������擾
    int_primaryKeyCnt = CInt(WorksheetFunction.CountIf(wkSheet.Range(P_PRIMARY_COLUMN & lng_stRow & ":" & P_PRIMARY_COLUMN & lng_edRow), "<>"))
    If int_primaryKeyCnt > 0 Then
        With primaryKeySet
            ReDim .ary_listSet(1 To int_primaryKeyCnt)
            .str_listName = Left(myTableInfo.str_TblName, 4) & "_PRIMARY_KEY"
        End With
    End If
    'INDEX �̓��͌������擾
    For int_loop = 1 To P_INDEX_COUNT
        int_totalIndexCnt(int_loop) = CInt(WorksheetFunction.CountIf(wkSheet.Range(P_INDEX1_COLUMN & lng_stRow & ":" & P_INDEX1_COLUMN & lng_edRow).Offset(0, int_loop - 1), "<>"))
        If int_totalIndexCnt(int_loop) > 0 Then
            With indexSet(int_loop)
                'INDEX NAME
                .str_listName = Left(myTableInfo.str_TblName, 4) & "_INDEX" & int_loop
                ReDim Preserve .ary_listSet(1 To int_totalIndexCnt(int_loop))
            End With
        End If
    Next int_loop
    

    
    '���̓f�[�^�͈̔͂��܂Ƃ߂Ĕz��Ɋi�[
    var_inputRange = wkSheet.Range(P_COLNAME_COLUMN & lng_stRow & ":" & P_INDEX3_COLUMN & lng_edRow)
    '������
    str_CreateDefinition = ""
    
    '�J��������1�s���m�F
    For lng_currentRow = lng_stRow To lng_edRow
        
        'No �J�E���g�A�b�v
        lng_No = (lng_currentRow - lng_stRow) + 1
        '�V�[�g�X�V
        With wkSheet
            'No ��������
            .Range(P_NO_COLUMN & lng_currentRow).Value = lng_No
        End With
        
        ' ���ݍs�̃J���������擾
        With myCurrentSet
            .str_ColumnName = var_inputRange(lng_No, 1)
            .str_ColumnComment = var_inputRange(lng_No, 2)
            .str_DataType = var_inputRange(lng_No, 3)
            .str_DataTypeDetail = var_inputRange(lng_No, 4) '�f�[�^�^�i�L���u(M)�v�u(M,D)�v���܂ށj
            If IsNumeric(var_inputRange(lng_No, 5)) Then
                .int_PRIMARY_KEY = Int(var_inputRange(lng_No, 5))
                If .int_PRIMARY_KEY >= 1 And .int_PRIMARY_KEY <= P_PRIMARY_MAX Then
                    With primaryKeySet.ary_listSet(.int_PRIMARY_KEY)
                        .int_listNo = myCurrentSet.int_PRIMARY_KEY
                        .str_ColumnName = myCurrentSet.str_ColumnName
                    End With
                End If
            Else
                .int_PRIMARY_KEY = 0
            End If
            .str_M = var_inputRange(lng_No, 6)
            .str_D = var_inputRange(lng_No, 7)
            .str_NOTNULL = var_inputRange(lng_No, 8)
            .str_DEFAULT = var_inputRange(lng_No, 9)
            .str_UNSIGNED = var_inputRange(lng_No, 10)
            .str_ZEROFILL = var_inputRange(lng_No, 11)
            ' INDEX 1 ~ 3
            ReDim .int_IndexNo(1 To P_INDEX_COUNT)
            For int_loop = 1 To P_INDEX_COUNT
                If IsNumeric(var_inputRange(lng_No, 11 + int_loop)) Then
                    .int_IndexNo(int_loop) = var_inputRange(lng_No, 11 + int_loop)
                    '
                    If .int_IndexNo(int_loop) > P_INDEX_MAX Then
                        errMsg = errMsg & "�C���f�b�N�X�̎w�肪����������܂���(1)�B" & vbNewLine
                        GoTo EXIT_CODE
                    ElseIf .int_IndexNo(int_loop) >= 1 And .int_IndexNo(int_loop) <= P_INDEX_MAX Then
                        With indexSet(int_loop)
                            'INDEX VALUE
                             .ary_listSet(myCurrentSet.int_IndexNo(int_loop)).int_listNo = myCurrentSet.int_IndexNo(int_loop)
                             .ary_listSet(myCurrentSet.int_IndexNo(int_loop)).str_ColumnName = myCurrentSet.str_ColumnName
                        End With
                    ElseIf .int_IndexNo(int_loop) = 0 Then
                        ' INDEX �Ȃ�
                    Else
                        errMsg = errMsg & "�C���f�b�N�X�̎w�肪����������܂���(2)�B" & vbNewLine
                        GoTo EXIT_CODE
                    End If
                Else
                    .int_IndexNo(int_loop) = 0
                End If
            Next int_loop
            

            '----------------------------------------------
            ' CREATE   column_definition
            '
            'column_definition:
            '    data_type -(1)
            '      [NOT NULL | NULL] [DEFAULT default_value] -(2)
            '      [AUTO_INCREMENT] [UNIQUE [KEY] | [PRIMARY] KEY] -(3)
            '      [COMMENT 'string'] -(4)
            '      [COLUMN_FORMAT {FIXED|DYNAMIC|DEFAULT}] -(5)
            '      [STORAGE {DISK|MEMORY|DEFAULT}] -(6)
            '      [reference_definition] -'(7)
           
            'column_definition(1) : data_type
            .str_DataTypeSQL = .str_DataTypeDetail
            .str_DataTypeSQL = Replace(.str_DataTypeSQL, "(M)", "(" & .str_M & ")", 1, , vbTextCompare) '(M) = (length)
            .str_DataTypeSQL = Replace(.str_DataTypeSQL, "(M,D)", "(" & .str_M & "," & .str_D & ")", 1, , vbTextCompare) '(M,D) = (length,decimals)
            Select Case .str_DataType
                Case P_TYPE01_NUMBER
                '���l�^ [UNSIGNED] [ZEROFILL]
                    If StrComp(Left(.str_DataTypeDetail, 4), "BIT", vbTextCompare) <> 0 Then
                        'ADD [UNSIGNED]
                        If StrComp(.str_UNSIGNED, "UNSIGNED", vbTextCompare) = 0 Then
                            .str_DataTypeSQL = .str_DataTypeSQL & " " & "UNSIGNED"
                        End If
                        ' ADD [ZEROFILL]
                        If StrComp(.str_ZEROFILL, "ZEROFILL", vbTextCompare) = 0 Then
                            .str_DataTypeSQL = .str_DataTypeSQL & " " & "ZEROFILL"
                        End If
                    End If
                Case P_TYPE02_STRING
                '������^
                    '[CHARACTER SET] �� �w��Ȃ� ���J�������ł͂Ȃ��A�e�[�u���S�̂Ŏw��
                    '[COLLATE] �� �w��Ȃ�
                Case P_TYPE03_DATETIME
                '���t�E���Ԍ^
                    '(fsp) �� �w��Ȃ�
                Case Else
                '�^�G���[
                errMsg = errMsg & "�^�̒�`������������܂���B" & lng_currentRow & "�s�ڂ��m�F���Ă��������B" & vbNewLine
                GoTo EXIT_CODE
            End Select
            
            'column_definition(2) : [NOT NULL | NULL] [DEFAULT default_value]
            If StrComp(.str_NOTNULL, "NOT NULL", vbTextCompare) = 0 Then ' [NOT NULL]
                .str_DataTypeSQL = .str_DataTypeSQL & " " & "NOT NULL"
            End If
            If .str_DEFAULT <> "" Then  '[DEFAULT]
                .str_DataTypeSQL = .str_DataTypeSQL & " " & "DEFAULT " & .str_DEFAULT
            End If
            '
            'column_definition(3) :
            '     [AUTO_INCREMENT]�F�����ł͎w�肵�Ȃ��i�f�[�^�^�̑I�����Ɏw��j
            '     [UNIQUE [KEY]     �F��Ή�
            '    | [PRIMARY] KEY]  �F��Ή� ���J�������ł͂Ȃ��A�e�[�u���S�̂Ŏw��
            '
            'column_definition(4) : [COMMENT 'string']
            If .str_ColumnComment <> "" Then
                .str_DataTypeSQL = .str_DataTypeSQL & " " & "COMMENT '" & .str_ColumnComment & "'"
            End If
            'column_definition(5) : [COLUMN_FORMAT {FIXED|DYNAMIC|DEFAULT}]�F��Ή�
            'column_definition(6) : [STORAGE {DISK|MEMORY|DEFAULT}]�F��Ή�
            'column_definition(7) : [reference_definition]�F��Ή�
        End With
        
        ' ###�J������`��SQL�������쐬
        If lng_currentRow > lng_stRow Then
            str_CreateDefinition = str_CreateDefinition & " , "
        Else
            str_CreateDefinition = str_CreateDefinition & " "
        End If
        'create_definition:
        '    col_name column_definition -(1)
        '  | [CONSTRAINT [symbol]] PRIMARY KEY [index_type] (index_col_name,...) -(2)
        '      [index_option] ...
        '  | {INDEX|KEY} [index_name] [index_type] (index_col_name,...) -(3)
        '      [index_option] ...
        '  | [CONSTRAINT [symbol]] UNIQUE [INDEX|KEY] -(4)
        '      [index_name] [index_type] (index_col_name,...)
        '      [index_option] ...
        '  | {FULLTEXT|SPATIAL} [INDEX|KEY] [index_name] (index_col_name,...) -(5)
        '      [index_option] ...
        '  | [CONSTRAINT [symbol]] FOREIGN KEY -(6)
        '      [index_name] (index_col_name,...) reference_definition
        '  | CHECK (expr) -(7)
        '
        With myCurrentSet
            ' create_definition(1) : { [ col_name ] [ column_definition ] }
            str_CreateDefinition = str_CreateDefinition & .str_ColumnName & " " & .str_DataTypeSQL & vbNewLine
        End With
    Next lng_currentRow
     'create_definition(2) : PRIMARY KEY
    If int_primaryKeyCnt > 0 Then
        str_CreateDefinition = str_CreateDefinition & " , PRIMARY KEY ("
        For int_loop = 1 To int_primaryKeyCnt
            With primaryKeySet
                If int_primaryKeyCnt > 1 Then
                    str_CreateDefinition = str_CreateDefinition & " , "
                End If
                str_CreateDefinition = str_CreateDefinition & .ary_listSet(int_loop).str_ColumnName
            End With
        Next
        str_CreateDefinition = str_CreateDefinition & ") " & vbNewLine
    End If
    'create_definition(3) : INDEX
    For int_loop = 1 To P_INDEX_COUNT
        If int_totalIndexCnt(int_loop) > 0 Then
            With indexSet(int_loop)
                ' {INDEX|KEY} [index_name] [index_type]
                str_CreateDefinition = str_CreateDefinition & " , INDEX " & .str_listName
                str_CreateDefinition = str_CreateDefinition & " ("
                For int_subLoop = 1 To int_totalIndexCnt(int_loop)
                    If int_subLoop > 1 Then
                        str_CreateDefinition = str_CreateDefinition & " , "
                    End If
                    ' (index_col_name,...)
                    str_CreateDefinition = str_CreateDefinition & .ary_listSet(int_subLoop).str_ColumnName
                Next int_subLoop
                str_CreateDefinition = str_CreateDefinition & ") " & vbNewLine
            End With
        Else
            Exit For
        End If
    Next int_loop
    

    '------ CRATE SQL START ----------------
    'CREATE [TEMPORARY] TABLE [IF NOT EXISTS] tbl_name -(A)
    '    (create_definition , ...) -(B)
    '    [table_options] -(C)
    '    [partition_options] -(D)
    '
    With myTableInfo
    ' (A) CREATE [TEMPORARY] TABLE [IF NOT EXISTS] tbl_name
        str_SQL = str_SQL & "CREATE"
        If .is_TEMPORARY_Table = True Then
            str_SQL = str_SQL & " TEMPORARY"
        End If
        str_SQL = str_SQL & " TABLE"
        str_SQL = str_SQL & " " & .str_DbName & "." & .str_TblName
        '(B) (create_definition , ...)
        str_SQL = str_SQL & " (" & vbNewLine
        str_SQL = str_SQL & str_CreateDefinition
        str_SQL = str_SQL & ")" & vbNewLine
        '(C) [table_options]
        'table_option:
        '    ENGINE [=] engine_name -(1)
        '  | AUTO_INCREMENT [=] value -(2)
        '  | AVG_ROW_LENGTH [=] value -(3)
        '  | [DEFAULT] CHARACTER SET [=] charset_name -(4)
        '  | CHECKSUM [=] {0 | 1} -(5)
        '  | [DEFAULT] COLLATE [=] collation_name -(6)
        '  | COMMENT [=] 'string'  -(7)
        '  | CONNECTION [=] 'connect_string' -(8)
        '  | DATA DIRECTORY [=] 'absolute path to directory' -(9)
        '  | DELAY_KEY_WRITE [=] {0 | 1} -(10)
        '  | INDEX DIRECTORY [=] 'absolute path to directory' -(11)
        '  | INSERT_METHOD [=] { NO | FIRST | LAST } -(12)
        '  | KEY_BLOCK_SIZE [=] value -(13)
        '  | MAX_ROWS [=] value -(14)
        '  | MIN_ROWS [=] value -(15)
        '  | PACK_KEYS [=] {0 | 1 | DEFAULT} -(16)
        '  | PASSWORD [=] 'string' -(17)
        '  | ROW_FORMAT [=] {DEFAULT|DYNAMIC|FIXED|COMPRESSED|REDUNDANT|COMPACT} -(18)
        '  | STATS_AUTO_RECALC [=] {DEFAULT|0|1} -(19)
        '  | STATS_PERSISTENT [=] {DEFAULT|0|1} -(20)
        '  | STATS_SAMPLE_PAGES [=] value -(21)
        '  | TABLESPACE tablespace_name [STORAGE {DISK|MEMORY|DEFAULT}] -(22)
        '  | UNION [=] (tbl_name[,tbl_name]...) -(23)
        '
        'table_option(1) : ENGINE [=] engine_name
        str_SQL = str_SQL & " ENGINE = " & .str_ENGINE & "" & vbNewLine
        'table_option(4) : [DEFAULT] CHARACTER SET [=] charset_name
        str_SQL = str_SQL & " CHARACTER SET = " & .str_CHARSET & "" & vbNewLine
        'table_option(7) : COMMENT [=] 'string'
        str_SQL = str_SQL & " COMMENT = '" & .str_TblComment & "'" & vbNewLine
         'table_option : ��L(1)(4)(7)�ȊO�͔�Ή�
        '
        '(D) [partition_options] : ��Ή�
    End With
    '
    str_SQL = str_SQL & ";" & vbNewLine
     '--- CRATE SQL END ------------------------
     
    
EXIT_CODE:
    Return_SQL_CREATE_TABLE = str_SQL
End Function


'===========================
'SQL�����V�[�g�ɏ�������
'����(1)�F�ΏۃV�[�g
'����(2)�F���s��SQL��
'����(3)�FSQL���o�͌��ʃZ�b�g
'����(4)�F�G���[���b�Z�[�W
'�߂�l�F���s���ʁi�����FTrue�j
'===========================
Public Function IsOK_WriteSQL(ByRef wkSheet As Worksheet _
                                            , ByVal str_SQL As String _
                                            , ByRef resSet As RESULT_SET _
                                            , ByRef errMsg As String) As Boolean
    Dim isOK As Boolean '�o�͌���
    '�\��t���f�[�^�p
    Dim lng_stRow As Long '�J�n�s
    Dim lng_edRow As Long '�I���s

 On Error GoTo EXIT_CODE

    With resSet
              
        '���s��؂�ňꎟ���z��Ɋi�[
        .strAry_SQL = Split(str_SQL, vbNewLine)
        
        lng_stRow = P_SQL_START_ROW '�J�n�s
        lng_edRow = lng_stRow + UBound(.strAry_SQL) '�I���s
        
        '�񎟌��z��ɕϊ���A�z����Z���͈͂ɂ܂Ƃ߂ď�������
        wkSheet.Range(P_SQL_COLUMN & lng_stRow & ":" & P_SQL_COLUMN & lng_edRow) = WorksheetFunction.Transpose(.strAry_SQL)
    End With
    
EXIT_CODE:
    If Err.Description <> "" Then
        errMsg = errMsg & "SQL�̌��ʏo�͂Ɏ��s���܂����B" & vbNewLine
        errMsg = errMsg & Err.Description & vbNewLine
        isOK = False
    Else
        errMsg = ""
        isOK = True
    End If

    IsOK_WriteSQL = isOK
End Function


