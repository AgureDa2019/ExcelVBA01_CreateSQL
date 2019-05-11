Attribute VB_Name = "Mod02_SUB"
Option Explicit

'==========================================
'テーブル情報の取得
'引数(1)：　対象シート (IN)
'引数(2)：　テーブル情報の取得データ (OUT)
'引数(3)：　エラーメッセージ (OUT)
'戻り値　結果（OK：True、NG：False）
'==========================================
Public Function Get_TableInfo(ByRef wkSheet As Worksheet _
                                         , ByRef myTableInfo As TABLE_INFO _
                                         , Optional errMsg As String) As Boolean
Dim isOK As Boolean

On Error GoTo ERR_CODE
    
    With myTableInfo
        ' D列
        .str_DbName = wkSheet.Range(P_DBNAME_CELL).Value
        .str_TblName = wkSheet.Range(P_TABLENAME_CELL).Value
        .str_TblComment = wkSheet.Range(P_TABLECOMMENT_CELL).Value
        ' F列
        .str_TblType = wkSheet.Range(P_TABLETYPE_CELL).Value
        .str_MySql_ver = wkSheet.Range(P_MYSQLVER_CELL).Value
        .str_ENGINE = wkSheet.Range(P_ENGINE_CELL).Value
        .str_CHARSET = wkSheet.Range(P_CHARSET_CELL).Value
        '一時テーブルかどうか
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
'実行結果のクリア
'引数(1)：　対象シート
'引数(2)：　実行結果の入力開始行
'引数(3)：　実行結果のシート最下行
'戻り値　なし
'==========================================
Public Sub Clear_ResultCells(ByRef mySheet As Worksheet _
                                        , ByRef lng_stRow As Long _
                                        , ByRef lng_edRow As Long)
    '--------------------------
    'シート内のセル値クリア
    '--------------------------
    With mySheet
    
    
        'No
        .Range(P_NO_COLUMN & P_DATA_START_ROW & ":" & P_NO_COLUMN & lng_edRow).Value = ""
        '出力結果
        .Range(P_RESULT_RENGE).Value = "" '処理結果
        .Range(P_SQL_COLUMN & lng_stRow & ":" & P_SQL_COLUMN & lng_edRow).Value = "" 'SQL文
        
        '-------------------------------
        '背景色だけをクリア: テーブル情報
        '-------------------------------
        'スキーマ名 *
        '物理テーブル名 *
        '論理テーブル名 *
        .Range(P_DBNAME_CELL _
                   & ":" & P_TABLECOMMENT_CELL).Interior.ColorIndex = 0 '背景色だけをクリア
        '利用目的 *
        .Range(P_PURPOSE_CELL).Interior.ColorIndex = 0 '背景色だけをクリア
        'テーブル区分 *
        'MySQL ver *
        'ENGINE *
        'CHARSET *
        .Range(P_TABLETYPE_CELL _
                   & ":" & P_CHARSET_CELL).Interior.ColorIndex = 0 '背景色だけをクリア
        
        '-------------------------------
        '背景色だけをクリア: カラム情報
        '-------------------------------
        .Range(P_PRIMARY_COLUMN & P_DATA_START_ROW _
                   & ":" & P_INDEX3_COLUMN & P_DATA_MAX_ROW).Interior.ColorIndex = 0 '背景色だけをクリア
    End With
    
End Sub

'==========================================
'テーブル定義　入力値チェック    ## (unfinished) ##
'引数(1)：　対象シート
'引数(2)：　カラム情報の入力開始行
'引数(3)：　カラム情報の入力最下行
'引数(4)：　テーブル情報(IN)
'引数(5)：　エラーメッセージ(OUT)
'戻り値　チェック結果（OK：True、NG：False）
'==========================================
Public Function IsOK_ValidationCheck(ByRef mySheet As Worksheet _
                                                     , ByRef lng_stRow As Long _
                                                     , ByRef lng_edRow As Long _
                                                     , ByRef myTableInfo As TABLE_INFO _
                                                     , ByRef errMsg As String) As Boolean
    Dim isOK                 As Boolean   'チェック結果
    Dim lng_TotalCnt     As Long         '指定したセル範囲内の入力件数
    Dim lng_OutOfRengeCnt As Long   '範囲外の入力件数
    'チェック行
    Dim lng_nowRow As Long 'カラム情報の入力チェック行
    Dim str_targetCell As String  'チェック対象セル
    'セル範囲
    Dim str_colNameRange As String 'カラム情報の入力セル範囲
    Dim str_OutOfRenge As String     'カラム情報の入力セル範囲外

    errMsg = ""
    
    On Error GoTo EXIT_CODE
    
    '-----------------------------------------------
    '(1/5) シートの形式チェック ※タイトルが正しいかチェック
    '-----------------------------------------------
    If mySheet.Range(P_LABEL_TABLE_CELL).Value <> "テーブル情報" _
     Or mySheet.Range(P_LABEL_COLUMN_CELL).Value <> "カラム情報" _
     Or mySheet.Range(P_NO_COLUMN & P_DATA_HEAD_ROW).Value <> "No" Then
        errMsg = errMsg & "シートのフォーマットが正しくありません。※雛形シートをコピーして再実行してください。"
        GoTo EXIT_CODE
    End If

    '-----------------------------------------------
    '(2/5) カラム情報の入力件数チェック
    '-----------------------------------------------
    str_colNameRange = P_COLNAME_COLUMN & P_DATA_START_ROW & ":" & P_COLNAME_COLUMN & P_DATA_MAX_ROW
    With mySheet
        lng_TotalCnt = WorksheetFunction.CountA(.Range(str_colNameRange))     'セル範囲内の入力件数
        lng_edRow = .Cells(Rows.Count, .Range(P_COLNAME_COLUMN & "1").Column).End(xlUp).Row          '連続した3列目（物理名列）の最下行
    End With
    '(2/5) - 1
    If lng_TotalCnt < 1 Then
        errMsg = errMsg & "カラム情報の[Column Name]列が未入力です。" & "チェック範囲：" & str_colNameRange & vbNewLine
    End If
    '(2/5) - 2
    If (lng_TotalCnt + P_DATA_HEAD_ROW) <> lng_edRow Then
        errMsg = errMsg & "カラム情報の[Column Name]は連続して入力してください。" & "チェック範囲：" & str_colNameRange & vbNewLine
    End If
    '(2/5) - 3
    If lng_edRow < P_DATA_MAX_ROW Then
        str_OutOfRenge = P_COLNAME_COLUMN & (lng_edRow + 1) & ":" & P_INDEX3_COLUMN & P_DATA_MAX_ROW
        lng_OutOfRengeCnt = WorksheetFunction.CountA(mySheet.Range(str_OutOfRenge))  'セル範囲外の入力件数
        If lng_OutOfRengeCnt > 0 Then
            errMsg = errMsg & "カラム情報に不要なデータが含まれています。" & "チェック範囲：" & str_OutOfRenge & vbNewLine
        End If
    End If
    '(2/5) エラーのため作業を中止
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    
    '-------------------------------------
    '(3/5) テーブル情報のチェック
    '--------------------------------------
    '---------------------------------
    ' (3/5)-1 テーブル情報 必須入力
    '---------------------------------
    '必須(1):[テーブル区分] *
    Call IsOK_InputRange(mySheet, P_TABLETYPE_CELL, "テーブル情報 の[テーブル区分]", errMsg)
    '必須(2):[スキーマ名] *
    Call IsOK_InputRange(mySheet, P_DBNAME_CELL, "テーブル情報 の[スキーマ名]", errMsg)
    '必須(3):[物理テーブル名] *
    Call IsOK_InputRange(mySheet, P_TABLENAME_CELL, "テーブル情報 の[物理テーブル名]", errMsg)
    '必須(4):[論理テーブル名] *
    Call IsOK_InputRange(mySheet, P_TABLECOMMENT_CELL, "テーブル情報 の[論理テーブル名]", errMsg)
    '必須(5):[MySQL ver] *
    Call IsOK_InputRange(mySheet, P_MYSQLVER_CELL, "テーブル情報 の[MySQL ver]", errMsg)
    '必須(6):[ENGINE] *
    Call IsOK_InputRange(mySheet, P_ENGINE_CELL, "テーブル情報 の[ENGINE]", errMsg)
     '必須(7):[CHARSET] *
    Call IsOK_InputRange(mySheet, P_CHARSET_CELL, "テーブル情報 の[CHARSET]", errMsg)
     '必須(8):[利用目的] *
    Call IsOK_InputRange(mySheet, P_PURPOSE_CELL, "テーブル情報 の[利用目的]", errMsg)
    '(3/5) エラーのため作業を中止
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    
    '-----------------------------------------------
    '(4/5)  カラム情報のチェック
    '-----------------------------------------------
    For lng_nowRow = lng_stRow To lng_edRow
        With mySheet
            '-----------------------------------------------
            ' (4/5)-1 カラム情報 必須入力
            '-----------------------------------------------
            '必須(1):[物理カラム名] *
            str_targetCell = P_COLNAME_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の[物理カラム名]", errMsg)
            '必須(1):[論理カラム名] *
            str_targetCell = P_COLCOMMENT_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の[論理カラム名]", errMsg)
            '必須(1):[型分類] *
            str_targetCell = P_TYPE_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の[型分類]", errMsg)
            '必須(1):[データ型] *
            str_targetCell = P_TYPEDETAIL_COLUMN & lng_nowRow
            Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の[データ型]", errMsg)
            
            '-----------------------------------------------
            '(4/5)-2 型チェック
            '-----------------------------------------------
            ' --- 数値型 / 文字列型---
            Select Case True
                Case CStr(.Range(P_TYPEDETAIL_COLUMN & lng_nowRow).Value) Like "*(M)*"
                'データ型に(M)を含む場合、(M)は入力必須
                    str_targetCell = P_M_COLUMN & lng_nowRow
                    Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の(M)", errMsg)
                Case CStr(.Range(P_TYPEDETAIL_COLUMN & lng_nowRow).Value) Like "*(M,D)*"
                'データ型に(M,D)を含む場合、(M)(D)は入力必須
                    str_targetCell = P_M_COLUMN & lng_nowRow
                    Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の(M)", errMsg)
                    str_targetCell = P_D_COLUMN & lng_nowRow
                    Call IsOK_InputRange(mySheet, str_targetCell, "カラム情報 の(D)", errMsg)
            End Select
            ' --- 日付・時間型 ---
            ' (tsp)は対象外
                  
            '-----------------------------------------------
            '(4/5)-3 長さチェック ## (unfinished) ##
            '-----------------------------------------------
            ' --- 数値型 ---
            ' --- 日付・時間型 ---
            ' --- 文字列型 ---
        End With
    Next lng_nowRow
    
    '(4/5) エラーのため作業を中止
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    
    '-----------------------------------------------
    '(5/5) MYSQL Version 8.0 CHECK
    '-----------------------------------------------
    With myTableInfo
        If .str_MySql_ver Like "*8.0*" _
            And StrComp(.str_CHARSET, "utf8", vbTextCompare) = 0 Then
            ' バージョンが8で文字コードがutf8の場合、utf8mb4が選択されていないとエラー扱い
            errMsg = errMsg & "MYSQL のバージョンが8で文字コードがutf8の場合、utf8mb4を選択してください。" & "チェック範囲：" & P_PURPOSE_CELL & vbNewLine
        End If
    End With
    
    '(5/5) エラーのため作業を中止
    If errMsg <> "" Then
        GoTo EXIT_CODE
    End If
    

     isOK = True
     
EXIT_CODE:
    IsOK_ValidationCheck = isOK
End Function

'====================================
'必須項目の入力チェックと、エラー時は対象セルを指定色で塗りつぶす
'引数(1)：対象シート
'引数(2)：対象セル名
'引数(3)：エラー時に出力する項目名
'引数(4)：エラーメッセージ
'戻り値：チェック結果
'====================================
Public Function IsOK_InputRange(ByRef mySheet As Worksheet _
                                                , ByRef str_CellName As String _
                                                , ByRef str_ItemName As String _
                                                , ByRef errMsg As String) As Boolean
    Dim isOK As Boolean
    Dim lng_targetColor As Long
    
    lng_targetColor = 44
    
    If Len(Trim(mySheet.Range(str_CellName).Value)) = 0 Then
    '必須項目のチェックNG
        errMsg = errMsg & str_ItemName & "が未入力です。" & "チェック範囲：" & str_CellName & "セル" & vbNewLine
        '指定色で塗りつぶす
        Call Change_RangeBGColor(mySheet.Range(str_CellName), lng_targetColor)
        isOK = False
    Else
    'チェックOK
        isOK = True
    End If
    
    IsOK_InputRange = isOK
End Function


'==========================================
'対象セル範囲の背景色を塗りつぶす
'引数(1)： 対象RANGE
'引数(2)： 背景色のインデックス番号   ex) RED=3
'戻り値　なし
'==========================================
Public Sub Change_RangeBGColor(ByRef targetRange As Range _
                                            , ByRef lng_colorIndex As Long)
    targetRange.Interior.ColorIndex = lng_colorIndex
End Sub


'==========================================
'テーブル定義 CREATE SQL文の作成
'引数(1)：　対象シート(IN/OUT)
'引数(2)：　カラム情報の入力開始行(IN)
'引数(3)：　カラム情報の入力最下行(IN)
'引数(4)：　テーブル情報(IN)
'引数(5)：　エラーメッセージ(OUT)
'戻り値　チェック結果（OK：True、NG：False）
' * https://dev.mysql.com/doc/refman/5.6/ja/create-table.html
'==========================================
Public Function Return_SQL_CREATE_TABLE(ByRef wkSheet As Worksheet _
                                                               , ByVal lng_stRow As Long _
                                                               , ByVal lng_edRow As Long _
                                                               , ByRef myTableInfo As TABLE_INFO _
                                                               , ByRef errMsg As String) As String
    Dim str_SQL As String                 '作成したSQL文
    Dim lng_No As Long                  '現在行のカラムNo
    Dim lng_currentRow As Long      '現在の行番号
    Dim myCurrentSet As CURRENT_COLUMN_SET '現在行のカラム情報
    Dim str_CreateDefinition As String 'カラム定義のSQL
    Dim var_inputRange As Variant '入力範囲内データ（配列として一時保持）
    Dim indexSet()  As COLUMN_LIST_SET 'INDEX1~3
    Dim primaryKeySet  As COLUMN_LIST_SET '主キーPRIMARY Key
    Dim int_totalIndexCnt() As Integer 'インデックスの入力件数
    Dim int_tmpIndexCnt As Integer 'インデックスの入力件数
    Dim int_primaryKeyCnt As Integer   '主キーの入力件数
    Dim int_loop As Integer
    Dim int_subLoop As Integer
    
    '初期化
    str_SQL = ""
    str_CreateDefinition = ""
    ReDim int_totalIndexCnt(1 To P_INDEX_COUNT)
    For int_loop = 1 To P_INDEX_COUNT
        int_totalIndexCnt(int_loop) = 0
    Next int_loop
    int_primaryKeyCnt = 0
    ReDim indexSet(1 To P_INDEX_COUNT)
    
    On Error GoTo EXIT_CODE
    
    'カラム情報の入力開始行と最下行が正しい範囲内かチェック
    If lng_stRow > P_DATA_HEAD_ROW _
       And lng_stRow < lng_edRow _
       And lng_edRow <= P_DATA_MAX_ROW Then
      'ROW CHECK OK
    Else
     'ROW CHECK NG
        errMsg = errMsg & "[ERROR] CREATE SQL ROW CHECK NG"
        GoTo EXIT_CODE
    End If
    
    'PRIMARY KEY の入力件数を取得
    int_primaryKeyCnt = CInt(WorksheetFunction.CountIf(wkSheet.Range(P_PRIMARY_COLUMN & lng_stRow & ":" & P_PRIMARY_COLUMN & lng_edRow), "<>"))
    If int_primaryKeyCnt > 0 Then
        With primaryKeySet
            ReDim .ary_listSet(1 To int_primaryKeyCnt)
            .str_listName = Left(myTableInfo.str_TblName, 4) & "_PRIMARY_KEY"
        End With
    End If
    'INDEX の入力件数を取得
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
    

    
    '入力データの範囲をまとめて配列に格納
    var_inputRange = wkSheet.Range(P_COLNAME_COLUMN & lng_stRow & ":" & P_INDEX3_COLUMN & lng_edRow)
    '初期化
    str_CreateDefinition = ""
    
    'カラム情報を1行ずつ確認
    For lng_currentRow = lng_stRow To lng_edRow
        
        'No カウントアップ
        lng_No = (lng_currentRow - lng_stRow) + 1
        'シート更新
        With wkSheet
            'No 書き込み
            .Range(P_NO_COLUMN & lng_currentRow).Value = lng_No
        End With
        
        ' 現在行のカラム情報を取得
        With myCurrentSet
            .str_ColumnName = var_inputRange(lng_No, 1)
            .str_ColumnComment = var_inputRange(lng_No, 2)
            .str_DataType = var_inputRange(lng_No, 3)
            .str_DataTypeDetail = var_inputRange(lng_No, 4) 'データ型（記号「(M)」「(M,D)」を含む）
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
                        errMsg = errMsg & "インデックスの指定が正しくありません(1)。" & vbNewLine
                        GoTo EXIT_CODE
                    ElseIf .int_IndexNo(int_loop) >= 1 And .int_IndexNo(int_loop) <= P_INDEX_MAX Then
                        With indexSet(int_loop)
                            'INDEX VALUE
                             .ary_listSet(myCurrentSet.int_IndexNo(int_loop)).int_listNo = myCurrentSet.int_IndexNo(int_loop)
                             .ary_listSet(myCurrentSet.int_IndexNo(int_loop)).str_ColumnName = myCurrentSet.str_ColumnName
                        End With
                    ElseIf .int_IndexNo(int_loop) = 0 Then
                        ' INDEX なし
                    Else
                        errMsg = errMsg & "インデックスの指定が正しくありません(2)。" & vbNewLine
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
                '数値型 [UNSIGNED] [ZEROFILL]
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
                '文字列型
                    '[CHARACTER SET] → 指定なし ※カラム毎ではなく、テーブル全体で指定
                    '[COLLATE] → 指定なし
                Case P_TYPE03_DATETIME
                '日付・時間型
                    '(fsp) → 指定なし
                Case Else
                '型エラー
                errMsg = errMsg & "型の定義が正しくありません。" & lng_currentRow & "行目を確認してください。" & vbNewLine
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
            '     [AUTO_INCREMENT]：ここでは指定しない（データ型の選択時に指定）
            '     [UNIQUE [KEY]     ：非対応
            '    | [PRIMARY] KEY]  ：非対応 ※カラム毎ではなく、テーブル全体で指定
            '
            'column_definition(4) : [COMMENT 'string']
            If .str_ColumnComment <> "" Then
                .str_DataTypeSQL = .str_DataTypeSQL & " " & "COMMENT '" & .str_ColumnComment & "'"
            End If
            'column_definition(5) : [COLUMN_FORMAT {FIXED|DYNAMIC|DEFAULT}]：非対応
            'column_definition(6) : [STORAGE {DISK|MEMORY|DEFAULT}]：非対応
            'column_definition(7) : [reference_definition]：非対応
        End With
        
        ' ###カラム定義のSQL部分を作成
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
         'table_option : 上記(1)(4)(7)以外は非対応
        '
        '(D) [partition_options] : 非対応
    End With
    '
    str_SQL = str_SQL & ";" & vbNewLine
     '--- CRATE SQL END ------------------------
     
    
EXIT_CODE:
    Return_SQL_CREATE_TABLE = str_SQL
End Function


'===========================
'SQL文をシートに書き込む
'引数(1)：対象シート
'引数(2)：改行つきSQL文
'引数(3)：SQL文出力結果セット
'引数(4)：エラーメッセージ
'戻り値：実行結果（成功：True）
'===========================
Public Function IsOK_WriteSQL(ByRef wkSheet As Worksheet _
                                            , ByVal str_SQL As String _
                                            , ByRef resSet As RESULT_SET _
                                            , ByRef errMsg As String) As Boolean
    Dim isOK As Boolean '出力結果
    '貼り付けデータ用
    Dim lng_stRow As Long '開始行
    Dim lng_edRow As Long '終了行

 On Error GoTo EXIT_CODE

    With resSet
              
        '改行区切りで一次元配列に格納
        .strAry_SQL = Split(str_SQL, vbNewLine)
        
        lng_stRow = P_SQL_START_ROW '開始行
        lng_edRow = lng_stRow + UBound(.strAry_SQL) '終了行
        
        '二次元配列に変換後、配列をセル範囲にまとめて書き込む
        wkSheet.Range(P_SQL_COLUMN & lng_stRow & ":" & P_SQL_COLUMN & lng_edRow) = WorksheetFunction.Transpose(.strAry_SQL)
    End With
    
EXIT_CODE:
    If Err.Description <> "" Then
        errMsg = errMsg & "SQLの結果出力に失敗しました。" & vbNewLine
        errMsg = errMsg & Err.Description & vbNewLine
        isOK = False
    Else
        errMsg = ""
        isOK = True
    End If

    IsOK_WriteSQL = isOK
End Function


