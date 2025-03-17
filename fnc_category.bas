Attribute VB_Name = "fnc_category"
Option Compare Database

Public Function category_name(category_id As Integer) As String
'
' ŒŸõ•¶š—ñ—p‚Éo—Í‚·‚é•K—v‚ ‚è
'

'
' parent_id_above : ‘æ“ñŠK‘w‚Ìid /‘æˆêŠK‘w‚Ìid ‚ğ•Ô‚·
'

'
' ?fnc_category.category_name(22)    :Šò•Œs;2023”N“x;A
' ?fnc_category.category_name(20)    :2025”N“x;A
' ?fnc_category.category_name(17)    :‚»‚Ì‘¼
' ?fnc_category.category_name(21)    :2024”N“x;B;
'
'eŠK‘w
    Dim parent_id As Integer
    parent_id = DLookup("parent_id", "category", "category_id=" & category_id)
'
    Dim ret As String
    ret = ""
'
'‘æ“ñŠK‘w‚Ìid
    Dim parent_id_above As Integer
    parent_id_previous = 0
    
    '‘æOŠK‘w‚Ü‚Å‚µ‚©‚È‚¢
    Select Case DLookup("category", "category", "category_id=" & category_id)
    Case 1
        '•\¦‚·‚é•¶š—ñ
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
    Case 2
        '‘æ“ñŠK‘w‚Ìid Šm’è
        parent_id_above = parent_id
        
        '•\¦‚·‚é•¶š—ñ
        ret = DLookup("categoryname", "category", "category_id=" & category_id) & ";"
        ret = ret & DLookup("categoryname", "category", "category_id=" & parent_id) & ";"
    Case 3
        '‘æˆêŠK‘w‚Ìid Šm’è
        parent_id_above = DLookup("parent_id", "category", "category_id=" & parent_id)
        
        '•\¦‚·‚é•¶š—ñ
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
    'eŠK‘w
    Dim parent_id As Integer
    parent_id = DLookup("parent_id", "category", "category_id=" & category_id)
'
    Dim ret As String
    ret = ""
'
'‘æ“ñŠK‘w‚Ìid
    Dim parent_id_above As Integer
    parent_id_previous = 0
    
    '‘æOŠK‘w‚Ü‚Å‚µ‚©‚È‚¢
    Select Case DLookup("category", "category", "category_id=" & category_id)
    Case 1
        '•\¦‚·‚é•¶š—ñ
        ret = CStr(category_id) & ";"
    Case 2
        '‘æ“ñŠK‘w‚Ìid Šm’è
        parent_id_above = parent_id
        
        '•\¦‚·‚é•¶š—ñ
        ret = CStr(category_id) & ";"
        ret = ret & CStr(parent_id_above) & ";"
    Case 3
        '‘æˆêŠK‘w‚Ìid Šm’è
        parent_id_above = DLookup("parent_id", "category", "category_id=" & parent_id)
        
        '•\¦‚·‚é•¶š—ñ
        ret = CStr(category_id) & ";"
        ret = ret & DLookup("category_id", "category", "category_id=" & parent_id) & ";"
        ret = ret & DLookup("category_id", "category", "category_id=" & parent_id_above) & ";"
        
    End Select
    
    category_id_str = ret

End Function

Public Function category_id_one(category_id As Integer) As Integer
'
' ƒJƒeƒSƒŠid‚æ‚èA‘æˆêŠK‘w‚¾‚¯•Ô‚·(ƒ[ƒ‹‚·‚é‘ÎÛ‚Æ‚È‚é‚Ì‚Å)
'

End Function
Public Function category_name_view(category_id As Integer) As String
'
' •\¦—p (ƒJƒeƒSƒŠ‚²‚Æ‚ÉAƒXƒy[ƒX‚ğ4‚Â’Ç‰Á‚·‚é)
'
    
'
' ?fnc_category.category_name_view(22)
'
    Select Case DLookup("category", "category ", "category_id=" & category_id)
    Case 0 '‘S‚Ä
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
