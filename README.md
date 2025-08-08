Function Levenshtein(s1 As String, s2 As String) As Integer
    Dim i As Integer, j As Integer
    Dim l1 As Integer, l2 As Integer
    Dim dist() As Integer
    Dim cost As Integer

    l1 = Len(s1)
    l2 = Len(s2)

    ReDim dist(0 To l1, 0 To l2)

    For i = 0 To l1
        dist(i, 0) = i
    Next i

    For j = 0 To l2
        dist(0, j) = j
    Next j

    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            dist(i, j) = Application.Min( _
                dist(i - 1, j) + 1, _
                dist(i, j - 1) + 1, _
                dist(i - 1, j - 1) + cost)
        Next j
    Next i

    Levenshtein = dist(l1, l2)
End Function

Function LevenshteinSimilarity(s1 As String, s2 As String) As Double
    Dim dist As Integer
    Dim maxLen As Integer

    s1 = LCase(Trim(s1))
    s2 = LCase(Trim(s2))

    dist = Levenshtein(s1, s2)
    maxLen = Application.Max(Len(s1), Len(s2))

    If maxLen = 0 Then
        LevenshteinSimilarity = 1
    Else
        LevenshteinSimilarity = 1 - (dist / maxLen)
    End If
End Function

Sub FuzzyMatchTextToWord()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rowNum As Long
    Dim paragraph As String, targetWord As String
    Dim words() As String
    Dim i As Integer
    Dim sim As Double
    Dim maxSim As Double
    Dim matchFound As Boolean

    rowNum = 1
    Do While ws.Cells(rowNum, 1).Value <> ""
        paragraph = ws.Cells(rowNum, 1).Value
        targetWord = ws.Cells(rowNum, 2).Value
        matchFound = False
        maxSim = 0

        ' Clean punctuation and split into words
        paragraph = Replace(paragraph, ".", "")
        paragraph = Replace(paragraph, ",", "")
        paragraph = Replace(paragraph, ";", "")
        paragraph = Replace(paragraph, ":", "")
        paragraph = Replace(paragraph, "!", "")
        paragraph = Replace(paragraph, "?", "")
        paragraph = Replace(paragraph, vbCr, "")
        paragraph = Replace(paragraph, vbLf, "")

        words = Split(paragraph, " ")

        For i = LBound(words) To UBound(words)
            sim = LevenshteinSimilarity(words(i), targetWord)
            If sim > maxSim Then maxSim = sim
            If sim >= 0.8 Then
                matchFound = True
            End If
        Next i

        ws.Cells(rowNum, 3).Value = matchFound
        ws.Cells(rowNum, 4).Value = Format(maxSim * 100, "0.00") & "%"

        rowNum = rowNum + 1
    Loop

    MsgBox "Fuzzy comparison with similarity % complete!"
End Sub


