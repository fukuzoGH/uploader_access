Attribute VB_Name = "fnc_category"
Option Compare Database

Public Function category_name(category_id As Integer) As String
'
' 検索文字列用に出力する必要あり
'

'
' parent_id_above : 第二階層のid /第一階層のid を返す
'

'
' ?fnc_category.category_name(22)    :岐阜市;2023年度;A
' ?fnc_category.category_name(20)    :2025年度;A
' ?fnc_category.category_name(17)    :その他
' ?fnc_category.category_name(21)    :2024年度;B;
'
'親階層
    Dim parent_id As Integer
    parent_id = DLookup("parent_id", "category", "category_id=" & category_id)
'
    Dim ret As String
    ret = ""
'
'第二階層のid
    Dim parent_id_above As Integer
    parent_id_previous = 0
    
    '第三階層までしかない
    Select Case DLookup("category", "category", "category_id=" & category_id)
    Case 1
        '表示する文字列
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
    Case 2
        '第二階層のid 確定
        parent_id_above = parent_id
        
        '表示する文字列
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
        ret = ret & DLookup("categoryname", "category", "category_id=" & parent_id) & ";"
    Case 3
        '第一階層のid 確定
        parent_id_above = DLookup("parent_id", "category", "category_id=" & parent_id)
        
        '表示する文字列
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
    '親階層
    Dim parent_id As Integer
    parent_id = DLookup("parent_id", "category", "category_id=" & category_id)
'
    Dim ret As String
    ret = ""
'
'第二階層のid
    Dim parent_id_above As Integer
    parent_id_previous = 0
    
    '第三階層までしかない
    Select Case DLookup("category", "category", "category_id=" & category_id)
    Case 1
        '表示する文字列
        ret = CStr(category_id) & ";"
    Case 2
        '第二階層のid 確定
        parent_id_above = parent_id
        
        '表示する文字列
        ret = CStr(category_id) & ";"
        ret = ret & CStr(parent_id_above) & ";"
    Case 3
        '第一階層のid 確定
        parent_id_above = DLookup("parent_id", "category", "category_id=" & parent_id)
        
        '表示する文字列
        ret = CStr(category_id) & ";"
        ret = ret & DLookup("category_id", "category", "category_id=" & parent_id) & ";"
        ret = ret & DLookup("category_id", "category", "category_id=" & parent_id_above) & ";"
        
    End Select
    
    category_id_str = ret

End Function

Public Function category_id_one(category_id As Integer) As Integer
'
' カテゴリidより、第一階層だけ返す(メールする対象となるので)
'

End Function
Public Function category_name_view(category_id As Integer) As String
'
' 表示用 (カテゴリごとに、スペースを4つ追加する)
'
    
'
' ?fnc_category.category_name_view(22)
'
    Select Case DLookup("category", "category ", "category_id=" & category_id)
    Case 0 '全て
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
