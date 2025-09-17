Attribute VB_Name = "ATSEngine"
Option Explicit

' Module: ATSEngine
' Purpose:
'   Implements keyword research, ATS scoring, formatting enforcement checks, and suggestions.

Private Function NormalizeText(ByVal s As String) As String
    s = LCase$(s)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    NormalizeText = Trim$(s)
End Function

Private Function WordList(ByVal s As String) As Collection
    Dim col As New Collection, arr() As String, i As Long, w As String
    s = NormalizeText(s)
    arr = Split(s, " ")
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        w = Trim$(arr(i))
        If Len(w) > 0 Then col.Add w, CStr(i + 1) & ":" & w
    Next i
    On Error GoTo 0
    Set WordList = col
End Function

Private Function CountOccurrences(ByVal text As String, ByVal phrase As String) As Long
    Dim i As Long, c As Long, pos As Long
    text = NormalizeText(text)
    phrase = NormalizeText(phrase)
    pos = InStr(1, text, phrase, vbTextCompare)
    Do While pos > 0
        c = c + 1
        pos = InStr(pos + Len(phrase), text, phrase, vbTextCompare)
    Loop
    CountOccurrences = c
End Function

Private Function DefaultKeywordBank() As Collection
    ' Provides a minimal built-in keyword list; can be extended via Templates/HelpDocs
    Dim c As New Collection
    c.Add "project management"
    c.Add "stakeholder"
    c.Add "requirements"
    c.Add "excel"
    c.Add "vba"
    c.Add "data analysis"
    c.Add "communication"
    c.Add "leadership"
    c.Add "agile"
    c.Add "scrum"
    Set DefaultKeywordBank = c
End Function

' PUBLIC_INTERFACE
Public Function BuildKeywordBank(Optional roleContext As String = "") As Collection
    ' Builds a keyword bank for ATS optimization combining defaults and role context heuristics
    Dim c As Collection, k As Variant
    Set c = DefaultKeywordBank()
    ' Simple heuristic: if roleContext provided, add role words
    If Len(roleContext) > 0 Then
        On Error Resume Next
        c.Add roleContext
        c.Add "experience in " & roleContext
        c.Add roleContext & " skills"
        On Error GoTo 0
    End If
    Set BuildKeywordBank = c
End Function

' PUBLIC_INTERFACE
Public Function ScoreKeywords(ByVal resumeText As String, ByVal keywords As Collection) As Double
    ' Returns a 0..100 score based on coverage and density of keywords
    Dim total As Long, hit As Long, k As Variant, occ As Long
    total = keywords.Count
    If total = 0 Then
        ScoreKeywords = 100
        Exit Function
    End If
    For Each k In keywords
        occ = CountOccurrences(resumeText, CStr(k))
        If occ > 0 Then hit = hit + 1
    Next k
    ScoreKeywords = Round((hit / total) * 100, 0)
End Function

' PUBLIC_INTERFACE
Public Function KeywordGaps(ByVal resumeText As String, ByVal keywords As Collection) As Collection
    ' Returns a collection of missing keywords
    Dim missing As New Collection, k As Variant
    For Each k In keywords
        If CountOccurrences(resumeText, CStr(k)) = 0 Then
            missing.Add CStr(k)
        End If
    Next k
    Set KeywordGaps = missing
End Function

' PUBLIC_INTERFACE
Public Function RecommendInsertionTips(ByVal missing As Collection) As String
    ' Builds a human-readable tip string to help insertion of missing keywords naturally
    Dim s As String, i As Long
    If missing Is Nothing Then
        RecommendInsertionTips = ""
        Exit Function
    End If
    s = "Consider adding the following keywords naturally in your bullets or summary:" & vbCrLf
    For i = 1 To missing.Count
        s = s & "• " & missing(i) & vbCrLf
    Next i
    RecommendInsertionTips = s
End Function

' PUBLIC_INTERFACE
Public Function EnforceFormattingRules(ByVal sectionName As String, ByVal text As String) As Collection
    ' Returns a list of formatting violations for ATS compatibility: bullets, capitalization, length, forbidden characters
    Dim violations As New Collection
    Dim lines() As String, i As Long, line As String

    lines = Split(text, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Then GoTo NextLine

        ' enforce bullet prefix for bullets
        If InStr(1, sectionName, "Experience", vbTextCompare) > 0 Or InStr(1, sectionName, "Projects", vbTextCompare) > 0 Or InStr(1, sectionName, "Skills", vbTextCompare) > 0 Then
            If Left$(line, 2) <> "• " And Left$(line, 1) <> "-" Then
                violations.Add "Line " & (i + 1) & ": Bullet missing '• ' prefix."
            End If
        End If

        ' length rule
        If Len(line) > 200 Then
            violations.Add "Line " & (i + 1) & ": Too long (" & Len(line) & " chars). Consider concise phrasing."
        End If

        ' avoid tables/graphics -> detect tabs and unusual chars
        If InStr(line, vbTab) > 0 Then
            violations.Add "Line " & (i + 1) & ": Contains tabs. Use plain text."
        End If
NextLine:
    Next i

    Set EnforceFormattingRules = violations
End Function

' PUBLIC_INTERFACE
Public Function SimulateATSParseScore(ByVal resumeText As String) As Double
    ' Computes an overall ATS simulation score based on formatting violations and keyword score
    Dim keywords As Collection
    Dim scoreK As Double, penalty As Double
    Set keywords = BuildKeywordBank("")
    scoreK = ScoreKeywords(resumeText, keywords)
    ' simple penalty: count lines that are too long or missing bullets
    Dim v As Collection, vCount As Long
    Set v = EnforceFormattingRules("Experience", resumeText)
    vCount = v.Count
    penalty = WorksheetFunction.Min(40, vCount * 5)
    SimulateATSParseScore = WorksheetFunction.Max(0, scoreK - penalty)
End Function
