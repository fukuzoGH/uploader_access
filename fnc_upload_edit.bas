Attribute VB_Name = "fnc_upload_edit"
Option Compare Database

Public Function categoryCnt(srtValue As String) As Integer
'
' 文字列;をカウントして何個か
'

'
' ?fnc_upload_edit.categoryCnt("何もない")  ':0
' ?fnc_upload_edit.categoryCnt("")          ':0
' ?fnc_upload_edit.categoryCnt("A;")        ':1
' ?fnc_upload_edit.categoryCnt("A;B;C;")    ':3
' ?fnc_upload_edit.categoryCnt("A;B;")      ':2
'
    categoryCnt = Len(srtValue) - Len(Replace(srtValue, ";", ""))
'
End Function
Public Function GetLeftPart(str As String) As String
'
' 文字列; で、左辺だけを取り出す
'

'
' ?fnc_upload_edit.GetLeftPart("A;")        ':A
' ?fnc_upload_edit.GetLeftPart("B;C;")      ':B
' ?fnc_upload_edit.GetLeftPart("C;B;A;")    ':C
'
'
    Dim pos As Long
    pos = InStr(str, ";")
    If pos > 0 Then
        GetLeftPart = Left(str, pos - 1)
    Else
        GetLeftPart = str
    End If
End Function
Function GetMiddleElement(str As String) As String
'
' 文字列; で、中央だけを取り出す
'

'
' ?fnc_upload_edit.GetMiddleElement("A;")       ':
' ?fnc_upload_edit.GetMiddleElement("B;C;")     ':
' ?fnc_upload_edit.GetMiddleElement("C;B;A;")   ':B
' ?fnc_upload_edit.GetMiddleElement("B;C;A;")   ':C
'
    If categoryCnt(str) <> 3 Then
            GetMiddleElement = ""
        Exit Function
    End If
'
    Dim elements() As String
    elements = Split(str, ";")
'
    Dim count As Integer
    count = UBound(elements) + 1
'
    If count Mod 2 = 0 Then
        GetMiddleElement = elements(count \ 2 - 1)
    Else
        GetMiddleElement = elements(count \ 2)
    End If
End Function
Function GetRightPart(str As String) As String
'
' 文字列; で、右辺だけを取り出す
'

'
' ?fnc_upload_edit.GetRightPart("A;")           ':
' ?fnc_upload_edit.GetRightPart("B;C;")         ':C
' ?fnc_upload_edit.GetRightPart("C;B;A;")       ':A
' ?fnc_upload_edit.GetRightPart("B;C;A;")       ':A
' ?fnc_upload_edit.GetRightPart("B-2;C-1;A-3;") ':A-3
'

    If categoryCnt(str) < 2 Then Exit Function
    
    Dim target  As String: target = ""
    If Mid(str, Len(str), 1) = ";" Then '最後の文字列に;があるとき
        target = Left(str, Len(str) - 1) ';をのぞく
    End If

    Dim ret As String: ret = ""
    Dim i As Integer
    For i = Len(target) To 1 Step -1
        If Mid(target, i, 1) = ";" Then
            GetRightPart = ret
            Exit For
        End If
        ret = Mid(target, i, 1) & ret
    Next i
    
End Function
