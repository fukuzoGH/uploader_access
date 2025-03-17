Attribute VB_Name = "fnc_upload"
Option Compare Database

Public Function url_hash(TextValue As String) As String
'
' ?fnc_upload.url_hash("設計書.pdf")
' ?fnc_upload.url_hash("図面1.pdf")
' ?fnc_upload.url_hash("設計書2.pdf")
' ?fnc_upload.url_hash("test")
' ?fnc_upload.url_hash("設計書2.pdf")
'
'
'

'ファイルのハッシュ値を求める
    'url_hash = MD5Hash(TextValue)
    url_hash = SHA256Hash(TextValue)

End Function


Private Function MD5Hash(ByVal sText As String) As String
'
' ハッシュ値を求める
'
' MD5
'
    Dim oMD5 As Object
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    Dim bText() As Byte
    bText = sText
    
    Dim bHash() As Byte
    bHash = oMD5.ComputeHash_2(bText)
    
    Dim sHash As String
    sHash = ""
    For i = 0 To UBound(bHash)
        sHash = sHash & Right("0" & Hex(bHash(i)), 2)
    Next
    
    MD5Hash = sHash
End Function
Private Function SHA256Hash(ByVal inputString As String) As String
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
Function RandomHash() As String
'
' ランダムなハッシュ
'
    Dim i As Integer
    Dim randomBytes(0 To 31) As Byte
    Dim hexString As String
    
    ' Generate 32 random bytes
    For i = 0 To 31
        randomBytes(i) = Int(Rnd() * 256)
    Next i
    
    ' Convert bytes to hexadecimal string
    For i = 0 To 31
        hexString = hexString & Right("0" & Hex(randomBytes(i)), 2)
    Next i
    
    RandomHash = LCase(hexString)
End Function
