Attribute VB_Name = "fnc_user"
Option Compare Database

Public Function SHA256Hash(ByVal inputString As String) As String
'
' ハッシュ値を求める (複雑で、比較的安全)
'
' SHA-256
'
    Dim hash As Object
    Dim bytes() As Byte
    Dim hashValue() As Byte
    Dim i As Integer
    Dim result As String
    
    Set hash = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    bytes = StrConv(inputString, vbFromUnicode)
    hashValue = hash.ComputeHash_2((bytes))
    
    For i = LBound(hashValue) To UBound(hashValue)
        result = result & LCase(Right("00" & Hex(hashValue(i)), 2))
    Next i
    
    SHA256Hash = result
End Function

Public Function category_list_user(user_id As Integer)
'
'ユーザ毎のカテゴリーidを準備 (タイミング:ログイン時)
'

'
' ?fnc_user.category_list_user(2)  'カテゴリー:15:A
' ?fnc_user.category_list_user(1)  'カテゴリー:27:(すべて)
'

Dim categoryList As New Collection
Dim categoryList1 As New Collection
Dim categoryList2 As New Collection
Dim categoryList3 As New Collection
'
Dim item As Variant
Dim item1 As Variant
Dim item2 As Variant
Dim item3 As Variant
'
Dim ca1 As DAO.Recordset
Set ca1 = CurrentDb.OpenRecordset("SELECT * FROM user_category1 WHERE user_id=" & user_id)
Do Until ca1.EOF
    categoryList1.Add ca1.Fields("category_id").value
    ca1.MoveNext
Loop
ca1.Close

'
Dim ca2 As DAO.Recordset
For Each item1 In categoryList1
    Set ca2 = CurrentDb.OpenRecordset("SELECT * FROM category WHERE parent_id=" & item1)
    Do Until ca2.EOF
        categoryList2.Add ca2.Fields("category_id").value
        ca2.MoveNext
    Loop
Next item1
ca2.Close
'

Dim ca3 As DAO.Recordset
For Each item2 In categoryList2
    If IsEmpty(item2) = False Then '何もカテゴリが無い場合もあるので
        Set ca3 = CurrentDb.OpenRecordset("SELECT * FROM category WHERE parent_id=" & item2)
        Do Until ca3.EOF
            categoryList3.Add ca3.Fields("category_id").value
            ca3.MoveNext
        Loop
        ca3.Close
    End If
Next item2


'
'
'

For Each item1 In categoryList1
    'Debug.Print item1
    categoryList.Add item1
Next item1
'
'Debug.Print "-----"
For Each item2 In categoryList2
    'Debug.Print item2
    categoryList.Add item2
Next item2
'
'Debug.Print "-----"
For Each item3 In categoryList3
    'Debug.Print item3
    categoryList.Add item3
Next item3
'
'
'
'
Dim rs As DAO.Recordset
CurrentDb.Execute "DELETE * FROM user_category WHERE user_id=" & user_id
For Each item In categoryList
    'Debug.Print item
    Set rs = CurrentDb.OpenRecordset("user_category")
    rs.AddNew
    rs.Fields("user_id").value = user_id
    rs.Fields("category_id").value = CInt(item)
    rs.Update
    rs.Close
Next item


End Function
