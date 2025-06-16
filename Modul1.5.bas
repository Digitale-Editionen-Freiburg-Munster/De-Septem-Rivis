Attribute VB_Name = "Modul1"
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
    Dim textLines() As String
    Dim line As Variant
    Dim processedLine As String
    
    ' Text aus dem Dokument holen
    textLines = Split(ActiveDocument.Content.text, vbCrLf)

    For Each line In textLines
        Dim wordList() As String
        wordList = Split(line, " ")
        processedLine = ""

        For Each w In wordList
            fullW = Replace(Replace(w, "(", ""), ")", "")
            
            ' Prüfe, ob das Wort eine Abkürzung enthält
            If fullW = w Then
                appendText = w
            Else
                abbrW = ""
                inBracket = False

                For i = 1 To Len(w)
                    c = Mid(w, i, 1)
                    If inBracket Then
                        If c = ")" Then inBracket = False
                    ElseIf c = "(" Then
                        inBracket = True
                    Else
                        abbrW = abbrW & c
                    End If
                Next i

                appendText = "<choice><abbr>" & abbrW & "</abbr><expan>" & fullW & "</expan></choice>"
            End If
            
            ' Wörter ohne unnötige Leerzeichen anfügen
            If processedLine = "" Then
                processedLine = appendText
            Else
                processedLine = processedLine & " " & appendText
            End If
        Next w
        
        ' Zeilenumbrüche beibehalten
        If newText = "" Then
            newText = processedLine
        Else
            newText = newText & vbCrLf & processedLine
        End If
    Next line

    ' Formatierten Text im Dokument einfügen
    ActiveDocument.Content.InsertAfter vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf
    ActiveDocument.Content.InsertAfter newText
End Sub

