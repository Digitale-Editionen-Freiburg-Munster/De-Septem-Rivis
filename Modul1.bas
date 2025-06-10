Attribute VB_Name = "Modul1"
Sub NewTest()
    ActiveDocument.Content.Delete
    ActiveDocument.Content.InsertAfter "T(e)st1 (Tes)t2 (Test3) T(e)s(t)4 Test5" '& vbCrLf & "Test6"
End Sub

Sub AbbreviationTEI()
    Dim newText As String
    Dim words() As String
    Dim fullW As String
    Dim abbrW As String
    Dim appendText As String
    Dim w As Variant
    Dim i As Integer
    Dim c As String
    Dim inBracket As Boolean

    text = ActiveDocument.Content.text
    words = Split(text, " ")

    For Each w In words
        fullW = Replace(Replace(w, "(", ""), ")", "")
        If fullW = w Then
            appendText = w
            GoTo append
        End If

        abbrW = ""
        inBracket = False

        For i = 1 To Len(w)
            c = Mid(w, i, 1)
            If inBracket Then
                If c = ")" Then inBracket = False
            ElseIf c = "(" Then
                inBracket = True
            Else: abbrW = abbrW & c
            End If
        Next i

        appendText = "<choice><abbr>" & abbrW & "</abbr><expan>" & fullW & "</expan></choice>"
append:
        newText = newText & appendText & " "
    Next w

    newText = RTrim(newText)
    ActiveDocument.Content.InsertAfter vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf
    ActiveDocument.Content.InsertAfter newText

End Sub
