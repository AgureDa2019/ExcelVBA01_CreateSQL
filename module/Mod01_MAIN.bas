Attribute VB_Name = "Mod01_Main"

'==========================================
' 【CREATE SQL】 BUTTON CLICK
'==========================================
Public Sub CreateSQL_ButtonClick()
    Dim isOK As Boolean                  '実行結果
    Dim errMsg As String                 'エラーメッセージ
    Dim mySheet As Worksheet        '対象シート
    Dim lng_stRow As Long               'カラム情報の入力開始行
    Dim lng_edRow    As Long         'カラム情報の入力最下行 ＊連続したセル範囲内（物理名列）
    Dim str_SQL As String                   'SQL
    Dim myTableInfo As TABLE_INFO  '入力されたテーブル情報
    Dim resSet As RESULT_SET 'SQL出力結果セット
    Dim str_msgText As String '実行結果（詳細）
    
    On Error GoTo EXIT_CODE
    
    Set mySheet = ActiveSheet 'アクティブシートを作業シートにセット
    lng_stRow = P_DATA_HEAD_ROW + 1  'カラム情報の入力開始行（＝見出し行＋1）
    
    '実行結果のクリア
    Call Clear_ResultCells(mySheet, P_SQL_START_ROW, CLng(Rows.Count))

    'シートに入力されたテーブル情報の取得
    isOK = Get_TableInfo(mySheet, myTableInfo)
    If isOK = False Then
        GoTo EXIT_CODE
    End If
    
    'セルの入力値チェックの実行
    isOK = IsOK_ValidationCheck(mySheet, lng_stRow, lng_edRow, myTableInfo, errMsg)
    If isOK = False Then
        GoTo EXIT_CODE
    End If
    
    ' CREATE TABLE SQL文の作成
    str_SQL = Return_SQL_CREATE_TABLE(mySheet, lng_stRow, lng_edRow, myTableInfo, errMsg)
    If errMsg <> "" Then
         isOK = False
    Else
        isOK = True
    End If
    
    'SQL文をシートに出力
    isOK = IsOK_WriteSQL(mySheet, str_SQL, resSet, errMsg)
    If isOK = False Then
        GoTo EXIT_CODE
    End If

EXIT_CODE:
    ' メッセージ作成 ----------------------
    Select Case isOK
        Case True
            str_msgText = "SQL文の作成に成功しました！"
        Case False
            str_msgText = "SQL文の作成に失敗しました！" & vbNewLine & errMsg
    End Select
    
    '実行結果の取得
    With resSet
        '処理時間
        .dat_resTime = Now()
        '実行結果
        .str_resultOK = IIf(isOK = True, "OK", "NG")
        '実行詳細
        .str_detailMsg = str_msgText
    End With
    
    '実行結果の書き込み
    With mySheet
        .Range(P_RESULT_DATE_CELL).Value = Format(resSet.dat_resTime, "YYYY/M/D HH:MM")
        .Range(P_RESULT_OKNG_CELL).Value = resSet.str_resultOK
        .Range(P_RESULT_DETAIL_CELL).Value = resSet.str_detailMsg
    End With
    
    ' メッセージ出力 ----------------------
    Select Case isOK
        Case True
        'OK
            Call Set_ScrollSheet(mySheet, "R1") 'SQL作成結果へスクロールを移動
            MsgBox str_msgText, vbInformation, "実行結果 [" & mySheet.Name & "]シート "
        Case False
        'NG
            Call Set_ScrollSheet(ActiveSheet, "A1") 'テーブル定義書入力へスクロールを移動
            MsgBox str_msgText, vbExclamation, "実行結果（エラー）[" & mySheet.Name & "]シート "
    End Select
    
    Set mySheet = Nothing
End Sub

'==========================================
' 【Back InputForm】 BUTTON CLICK
'==========================================
Public Sub BackInputForm_ButtonClick()
    'シート移動
    Call Set_ScrollSheet(ActiveSheet, "A1") 'テーブル定義書入力へスクロールを移動
End Sub

'==========================================
' 【COPY CREATESQL】 BUTTON CLICK
'==========================================
Public Sub CopyCreateSQL_ButtonClick()
    Dim isOK As Boolean                  '実行結果
    Dim errMsg As String                 'エラーメッセージ
    Dim mySheet As Worksheet        '対象シート
    Dim lng_stRow As Long               'SQLの出力開始行
    Dim lng_edRow    As Long         'SQLの出力最下行 ＊連続したセル範囲内（物理名列）

    On Error GoTo EXIT_CODE
    
    Set mySheet = ActiveSheet 'アクティブシートを作業シートにセット
    lng_stRow = P_SQL_START_ROW  'SQLの出力開始行
    With mySheet
        lng_edRow = .Cells(Rows.Count, .Range(P_SQL_COLUMN & "1").Column).End(xlUp).Row             '連続した対象列の最下行
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
