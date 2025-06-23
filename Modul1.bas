Attribute VB_Name = "Modul1"
Sub NewTest()
    ActiveDocument.Content.Delete
    ActiveDocument.Content.InsertAfter "T(e)st1 (Tes)t2 (Test3) T(e)s(t)4 Test5" & vbCr
    ActiveDocument.Content.InsertAfter "Test6 Te-" & vbCr & "st7 T(e)s-" & vbCr & "t8"
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
    Dim lineBreak As Boolean
    Dim wordBreak As Boolean

    text = ActiveDocument.Content.text
    words = Split(text, " ")

    For Each w In words
checkWord:
        fullW = Replace(w, "(", "")
        fullW = Replace(fullW, ")", "")

        If fullW = w Then
            ' Ob ein getrenntes oder zwei Woerter kann hier - weil ohne Abkuerzungen - ignoriert werden
            appendText = Replace(w, vbCr, "<lb/>")
            GoTo append
        End If

        fullW = Replace(fullW, vbCr, "<lb/>") ' Nur bei Worttrennung richtig

        abbrW = ""
        inBracket = False
        lineBreak = False ' Nur Line Break NACH einem Wort
        wordBreak = False

        For i = 1 To Len(w)
            c = Mid(w, i, 1)
            If inBracket Then
                If c = ")" Then inBracket = False
            ElseIf c = "(" Then
                inBracket = True

            ' Linke Break abarbeiten
            ElseIf c = vbCr Then
                If wordBreak Then ' fullW ist richtig, alles i.O.
                    abbrW = abbrW & "<lb/>"
                Else ' <lb/> darf nicht in choice-Tag
                    lineBreak = True
                    w = Split(fullW, "<lb/>")(1)
                    fullW = Split(fullW, "<lb/>")(0)
                    appendText = "<choice><abbr>" & abbrW & "</abbr><expan>" & fullW & "</expan></choice>"
                    newText = newText & appendText & "<lb/>"
                    GoTo checkWord
                End If

            Else
                abbrW = abbrW & c
                If c = "-" Then wordBreak = True
            End If
        Next i

        appendText = "<choice><abbr>" & abbrW & "</abbr><expan>" & fullW & "</expan></choice>"
append:
        newText = newText & appendText & " "
    Next w

    newText = RTrim(newText)
    ActiveDocument.Content.InsertAfter vbCr & vbCr & "---" & vbCr & vbCr
    ActiveDocument.Content.InsertAfter newText

End Sub
