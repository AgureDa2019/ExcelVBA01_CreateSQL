Attribute VB_Name = "Mod99_TEST"
Option Explicit

Public Sub TestCode()
    Dim int_osType As Integer '定義したOS識別子（1:Mac 2:Win 3:Other)
    Dim str_osVersion As String 'OSバージョン

    'OSバージョン情報を取得
    int_osType = Get_MyOS(str_osVersion)
    
    '確認（テストコード）
    Debug.Print int_osType
    Debug.Print str_osVersion
  
    Dim str_exlVer As String
    
    'EXCELのバージョンを取得
    str_exlVer = GET_MyExcelVersion
    
    '確認（テストコード）
    Debug.Print str_exlVer
    
    Dim str_lineBr As String '改行文字
    
    '改行文字を取得
    str_lineBr = GET_LineBreak
    
    '確認（テストコード）
    Debug.Print "aaa" & str_lineBr & "bbb"
End Sub


Sub export_all_module()
    Dim module_count As Long 'モジュールの個数
    Dim i As Long 'For文のカウンタとして使用
    
    With Application.VBE.ActiveVBProject.VBComponents
        module_count = .Count
        For i = 1 To module_count
            Select Case .Item(i).Type
                Case 3                                      'ユーザフォームのとき
                    .Item(i).Export (.Item(i).Name & ".frm")
                Case 1                                      '標準モジュールのとき
                    .Item(i).Export (.Item(i).Name & ".bas")
                Case Else                                   'それ以外
                    .Item(i).Export (.Item(i).Name & ".cls")
            End Select
        Next
    End With
End Sub
