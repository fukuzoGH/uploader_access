Attribute VB_Name = "fnc_category"
Option Compare Database

Public Function category_name(category_id As Integer) As String
'
' ����������p�ɏo�͂���K�v����
'

'
' parent_id_above : ���K�w��id /���K�w��id ��Ԃ�
'

'
' ?fnc_category.category_name(22)    :�򕌎s;2023�N�x;A
' ?fnc_category.category_name(20)    :2025�N�x;A
' ?fnc_category.category_name(17)    :���̑�
' ?fnc_category.category_name(21)    :2024�N�x;B;
'
'�e�K�w
    Dim parent_id As Integer
    parent_id = DLookup("parent_id", "category", "category_id=" & category_id)
'
    Dim ret As String
    ret = ""
'
'���K�w��id
    Dim parent_id_above As Integer
    parent_id_previous = 0
    
    '��O�K�w�܂ł����Ȃ�
    Select Case DLookup("category", "category", "category_id=" & category_id)
    Case 1
        '�\�����镶����
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
    Case 2
        '���K�w��id �m��
        parent_id_above = parent_id
        
        '�\�����镶����
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
        ret = ret & DLookup("categoryname", "category", "category_id=" & parent_id) & ";"
    Case 3
        '���K�w��id �m��
        parent_id_above = DLookup("parent_id", "category", "category_id=" & parent_id)
        
        '�\�����镶����
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
        ret = ret & DLookup("categoryname", "category", "category_id=" & parent_id) & ";"
        ret = ret & DLookup("categoryname", "category", "category_id=" & parent_id_above) & ";"
        
    End Select
    
    category_name = ret
    
End Function
Public Function category_id_str(category_id As Integer) As String
'
'
'

'
' ?fnc_category.category_id_str(22)    :22;18;15;
' ?fnc_category.category_id_str(20)    :20;15;
' ?fnc_category.category_id_str(17)    :17;
' ?fnc_category.category_id_str(21)    :21;16;
'
    '�e�K�w
    Dim parent_id As Integer
    parent_id = DLookup("parent_id", "category", "category_id=" & category_id)
'
    Dim ret As String
    ret = ""
'
'���K�w��id
    Dim parent_id_above As Integer
    parent_id_previous = 0
    
    '��O�K�w�܂ł����Ȃ�
    Select Case DLookup("category", "category", "category_id=" & category_id)
    Case 1
        '�\�����镶����
        ret = CStr(category_id) & ";"
    Case 2
        '���K�w��id �m��
        parent_id_above = parent_id
        
        '�\�����镶����
        ret = CStr(category_id) & ";"
        ret = ret & CStr(parent_id_above) & ";"
    Case 3
        '���K�w��id �m��
        parent_id_above = DLookup("parent_id", "category", "category_id=" & parent_id)
        
        '�\�����镶����
        ret = CStr(category_id) & ";"
        ret = ret & DLookup("category_id", "category", "category_id=" & parent_id) & ";"
        ret = ret & DLookup("category_id", "category", "category_id=" & parent_id_above) & ";"
        
    End Select
    
    category_id_str = ret

End Function

Public Function category_id_one(category_id As Integer) As Integer
'
' �J�e�S��id���A���K�w�����Ԃ�(���[������ΏۂƂȂ�̂�)
'

End Function
Public Function category_name_view(category_id As Integer) As String
'
' �\���p (�J�e�S�����ƂɁA�X�y�[�X��4�ǉ�����)
'
    
'
' ?fnc_category.category_name_view(22)
'
    Select Case DLookup("category", "category ", "category_id=" & category_id)
    Case 0 '�S��
        ret_category_name = DLookup("categoryname", "category", "category_id=" & category_id)
    Case 1
        ret_category_name = DLookup("categoryname", "category", "category_id=" & category_id)
    Case 2
        ret_category_name = Space(4) & DLookup("categoryname", "category", "category_id=" & category_id)
    Case 3
        ret_category_name = Space(8) & DLookup("categoryname", "category", "category_id=" & category_id)
    End Select
'
    category_name_view = ret_category_name
End Function
