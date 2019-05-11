Attribute VB_Name = "Mod00_DEFINE"
Option Explicit


'==========================================
'  CONST
'==========================================
'行番号
Public Const P_DATA_HEAD_ROW = 11
Public Const P_DATA_START_ROW = 12
Public Const P_DATA_MAX_ROW = 60
Public Const P_SQL_START_ROW = 11
'
Public Const P_INDEX_COUNT = 3
Public Const P_INDEX_MAX = 10
Public Const P_PRIMARY_MAX = 5
'
'型名
Public Const P_TYPE01_NUMBER = "数値型"
Public Const P_TYPE02_STRING = "文字列型"
Public Const P_TYPE03_DATETIME = "日付・時間型"
'列名
Public Const P_NO_COLUMN = "B"  'Noの列名
Public Const P_COLNAME_COLUMN = "C"       '物理カラム名 *
Public Const P_COLCOMMENT_COLUMN = "D" '論理カラム名 *
Public Const P_TYPE_COLUMN = "E"              '型分類 *
Public Const P_TYPEDETAIL_COLUMN = "F"   'データ型 *
Public Const P_PRIMARY_COLUMN = "G"       '主キー
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
'セル [テーブル情報]  *必須入力
Public Const P_DBNAME_CELL = "D5"                 ' スキーマ名 *(DATABASE NAME)
Public Const P_TABLENAME_CELL = "D6"            ' 物理テーブル名 *(TABLE NAME)
Public Const P_TABLECOMMENT_CELL = "D7"     ' 論理テーブル名 *(TABLE COMMENT)
Public Const P_TABLETYPE_CELL = "F4"             ' テーブル区分 * (TEMPORARY TABLE | TABLE)
Public Const P_MYSQLVER_CELL = "F5"              ' MySQL ver *
Public Const P_ENGINE_CELL = "F6"                   ' ENGINE *
Public Const P_CHARSET_CELL = "F7"                ' CHARSET *
Public Const P_PURPOSE_CELL = "D8"                ' 利用目的 * What is the purpose for this table?
'タイトル確認用
Public Const P_LABEL_TABLE_CELL = "B2"
Public Const P_LABEL_COLUMN_CELL = "B10"
'実行結果
Public Const P_RESULT_DATE_CELL = "T3"
Public Const P_RESULT_OKNG_CELL = "T4"
Public Const P_RESULT_DETAIL_CELL = "T5"

'セル範囲
Public Const P_RESULT_RENGE = "T3:T5"

'==========================================
'  TYPE
'==========================================

'テーブル情報
Public Type TABLE_INFO
    str_DbName As String           'スキーマ名
    str_TblName As String           '物理テーブル名
    str_TblComment As String     '論理テーブル名 (table comment)
    str_ENGINE  As String           'ENGINE
    str_CHARSET As String          'CHARSET
    str_MySql_ver As String         'MySQL ver MEMO
    str_TblType As String            'テーブル区分
    is_TEMPORARY_Table As Boolean '一時テーブルかどうか
    str_PURPOSE As String   '利用目的
'     'OPTION(NOT USED)
'    str_systemName As String
'    str_subSystemName As String
End Type


'カラム情報（1行分）
Public Type CURRENT_COLUMN_SET
        str_ColumnName As String       '物理カラム名
        str_ColumnComment As String  '論理カラム名
        str_DataType As String  'データ型（日本語表記）
        str_DataTypeDetail As String  'データ型（記号「(M)」「(M,D)」を含む）
        str_DataTypeSQL As String  'データ型（記号「(M)」「(M,D)」置換後でSQL出力用）
        str_M As String '記号「(M)」
        str_D As String '記号「(D)」
        str_NOTNULL As String
        str_DEFAULT As String
        str_UNSIGNED As String
        str_ZEROFILL As String
        '
        int_PRIMARY_KEY  As Integer
        int_IndexNo()  As Integer
End Type

'番号付きカラム一覧
Public Type COLUMN_LIST
  int_listNo As Integer                'NO
  str_ColumnName As String       '物理カラム名
End Type

'特定条件のカラム情報セット
Public Type COLUMN_LIST_SET
    str_listName As String                '情報名 ex) indexName, KeyName
    ary_listSet() As COLUMN_LIST    '番号付きカラム一覧
End Type

'SQL出力結果セット
Public Type RESULT_SET
    dat_resTime As Date     '処理時間
    str_resultOK As String  '実行結果
    str_detailMsg As String '実行詳細
    strAry_SQL() As String  'シート出力用SQL文（改行区切りの文字列を配列に格納）
End Type

