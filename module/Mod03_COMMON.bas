Attribute VB_Name = "Mod03_COMMON"
Option Explicit

'OS識別子
Public Enum OS_TYPE
    T01_MAC = 1
    T02_WIN = 2
    T03_OTHER = 0
End Enum

'文字コード
Public Const P_vbLf = 10 'ラインフィード文字
Public Const P_vbCr = 13 'キャリッジリターン文字
Public Const P_vbTab = 8 'タブ文字

'=================================
'Excel実行環境のOSを調べて返す
'引数(1)：取得したOSバージョン（省略可）
'戻り値：OS識別子（Enumで定義したもの）
'=================================
Public Function Get_MyOS(Optional ByRef str_ResOS As String) As Integer
    Dim str_OS As String 'OS
    Dim int_myOS As Integer

    'OSのバージョンを取得する
    str_OS = Application.OperatingSystem
    str_ResOS = str_OS

    'OS識別子をセット
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
'実行環境のExcelバージョンを調べて返す
'引数(1)：なし
'戻り値：OS識別子（Enumで定義したもの）
'=================================
Public Function GET_MyExcelVersion() As String
    Dim str_VER As String 'OS

    'OSのバージョンを取得する
    str_VER = Application.Version

    GET_MyExcelVersion = str_VER
End Function
 
 
'=================================
'実行環境のOS別の改行コードを調べて返す
'引数(1)：なし
'戻り値：改行文字
'=================================
Public Function GET_LineBreak() As String
    Dim str_LineBreak As String '改行

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
'アクティブセルを画面の左上端に表示する
'引数(1)：対象シート
'引数(2)：対象セル（左端に表示したいセル）
'引数(3)：エラーメッセージ（省略可）
'戻り値：なし
'==============================
Public Sub Set_ScrollSheet(ByRef wkSheet As Worksheet _
                                , ByVal tagetCell As String _
                                , Optional ByRef errMsg As String)
    On Error GoTo EXIT_CODE
    'シートをアクティブにして、指定されたセルを選択
    With wkSheet
        .Activate
        .Range(tagetCell).Select
    End With
    'アクティブセルへのスクロール移動
    With ActiveWindow
        .ScrollRow = ActiveCell.Row
        .ScrollColumn = ActiveCell.Column
    End With
EXIT_CODE:
    If Err.Description <> "" Then
        errMsg = errMsg & Err.Description & vbNewLine
    End If
End Sub
