Attribute VB_Name = "Mod00_DEFINE"
Option Explicit


'==========================================
'  CONST
'==========================================
'�s�ԍ�
Public Const P_DATA_HEAD_ROW = 11
Public Const P_DATA_START_ROW = 12
Public Const P_DATA_MAX_ROW = 60
Public Const P_SQL_START_ROW = 11
'
Public Const P_INDEX_COUNT = 3
Public Const P_INDEX_MAX = 10
Public Const P_PRIMARY_MAX = 5
'
'�^��
Public Const P_TYPE01_NUMBER = "���l�^"
Public Const P_TYPE02_STRING = "������^"
Public Const P_TYPE03_DATETIME = "���t�E���Ԍ^"
'��
Public Const P_NO_COLUMN = "B"  'No�̗�
Public Const P_COLNAME_COLUMN = "C"       '�����J������ *
Public Const P_COLCOMMENT_COLUMN = "D" '�_���J������ *
Public Const P_TYPE_COLUMN = "E"              '�^���� *
Public Const P_TYPEDETAIL_COLUMN = "F"   '�f�[�^�^ *
Public Const P_PRIMARY_COLUMN = "G"       '��L�[
Public Const P_M_COLUMN = "H"   '(M)
Public Const P_D_COLUMN = "I"     '(D)
Public Const P_NOTNULL_COLUMN = "J"
Public Const P_DEFAULT_COLUMN = "K"
Public Const P_UNSIGNED_COLUMN = "L"
Public Const P_ZEOFILL_COLUMN = "M"
Public Const P_INDEX1_COLUMN = "N"
Public Const P_INDEX2_COLUMN = "O"
Public Const P_INDEX3_COLUMN = "P"
Public Const P_SQL_COLUMN = "S"
'�Z�� [�e�[�u�����]  *�K�{����
Public Const P_DBNAME_CELL = "D5"                 ' �X�L�[�}�� *(DATABASE NAME)
Public Const P_TABLENAME_CELL = "D6"            ' �����e�[�u���� *(TABLE NAME)
Public Const P_TABLECOMMENT_CELL = "D7"     ' �_���e�[�u���� *(TABLE COMMENT)
Public Const P_TABLETYPE_CELL = "F4"             ' �e�[�u���敪 * (TEMPORARY TABLE | TABLE)
Public Const P_MYSQLVER_CELL = "F5"              ' MySQL ver *
Public Const P_ENGINE_CELL = "F6"                   ' ENGINE *
Public Const P_CHARSET_CELL = "F7"                ' CHARSET *
Public Const P_PURPOSE_CELL = "D8"                ' ���p�ړI * What is the purpose for this table?
'�^�C�g���m�F�p
Public Const P_LABEL_TABLE_CELL = "B2"
Public Const P_LABEL_COLUMN_CELL = "B10"
'���s����
Public Const P_RESULT_DATE_CELL = "T3"
Public Const P_RESULT_OKNG_CELL = "T4"
Public Const P_RESULT_DETAIL_CELL = "T5"

'�Z���͈�
Public Const P_RESULT_RENGE = "T3:T5"

'==========================================
'  TYPE
'==========================================

'�e�[�u�����
Public Type TABLE_INFO
    str_DbName As String           '�X�L�[�}��
    str_TblName As String           '�����e�[�u����
    str_TblComment As String     '�_���e�[�u���� (table comment)
    str_ENGINE  As String           'ENGINE
    str_CHARSET As String          'CHARSET
    str_MySql_ver As String         'MySQL ver MEMO
    str_TblType As String            '�e�[�u���敪
    is_TEMPORARY_Table As Boolean '�ꎞ�e�[�u�����ǂ���
    str_PURPOSE As String   '���p�ړI
'     'OPTION(NOT USED)
'    str_systemName As String
'    str_subSystemName As String
End Type


'�J�������i1�s���j
Public Type CURRENT_COLUMN_SET
        str_ColumnName As String       '�����J������
        str_ColumnComment As String  '�_���J������
        str_DataType As String  '�f�[�^�^�i���{��\�L�j
        str_DataTypeDetail As String  '�f�[�^�^�i�L���u(M)�v�u(M,D)�v���܂ށj
        str_DataTypeSQL As String  '�f�[�^�^�i�L���u(M)�v�u(M,D)�v�u�����SQL�o�͗p�j
        str_M As String '�L���u(M)�v
        str_D As String '�L���u(D)�v
        str_NOTNULL As String
        str_DEFAULT As String
        str_UNSIGNED As String
        str_ZEROFILL As String
        '
        int_PRIMARY_KEY  As Integer
        int_IndexNo()  As Integer
End Type

'�ԍ��t���J�����ꗗ
Public Type COLUMN_LIST
  int_listNo As Integer                'NO
  str_ColumnName As String       '�����J������
End Type

'��������̃J�������Z�b�g
Public Type COLUMN_LIST_SET
    str_listName As String                '��� ex) indexName, KeyName
    ary_listSet() As COLUMN_LIST    '�ԍ��t���J�����ꗗ
End Type

'SQL�o�͌��ʃZ�b�g
Public Type RESULT_SET
    dat_resTime As Date     '��������
    str_resultOK As String  '���s����
    str_detailMsg As String '���s�ڍ�
    strAry_SQL() As String  '�V�[�g�o�͗pSQL���i���s��؂�̕������z��Ɋi�[�j
End Type

