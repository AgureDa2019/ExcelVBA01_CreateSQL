Attribute VB_Name = "Mod99_TEST"
Option Explicit

Public Sub TestCode()
    Dim int_osType As Integer '��`����OS���ʎq�i1:Mac 2:Win 3:Other)
    Dim str_osVersion As String 'OS�o�[�W����

    'OS�o�[�W���������擾
    int_osType = Get_MyOS(str_osVersion)
    
    '�m�F�i�e�X�g�R�[�h�j
    Debug.Print int_osType
    Debug.Print str_osVersion
  
    Dim str_exlVer As String
    
    'EXCEL�̃o�[�W�������擾
    str_exlVer = GET_MyExcelVersion
    
    '�m�F�i�e�X�g�R�[�h�j
    Debug.Print str_exlVer
    
    Dim str_lineBr As String '���s����
    
    '���s�������擾
    str_lineBr = GET_LineBreak
    
    '�m�F�i�e�X�g�R�[�h�j
    Debug.Print "aaa" & str_lineBr & "bbb"
End Sub


Sub export_all_module()
    Dim module_count As Long '���W���[���̌�
    Dim i As Long 'For���̃J�E���^�Ƃ��Ďg�p
    
    With Application.VBE.ActiveVBProject.VBComponents
        module_count = .Count
        For i = 1 To module_count
            Select Case .Item(i).Type
                Case 3                                      '���[�U�t�H�[���̂Ƃ�
                    .Item(i).Export (.Item(i).Name & ".frm")
                Case 1                                      '�W�����W���[���̂Ƃ�
                    .Item(i).Export (.Item(i).Name & ".bas")
                Case Else                                   '����ȊO
                    .Item(i).Export (.Item(i).Name & ".cls")
            End Select
        Next
    End With
End Sub
