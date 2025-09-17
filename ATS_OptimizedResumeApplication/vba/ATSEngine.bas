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
    ' Builds a keyword bank combining:
    '  - built-in defaults (hardcoded)
    '  - Data! "keyword:default" (optional)
    '  - Data! "keyword:{role}" when roleContext specified, else DataModel.GetActiveRole
    '  - Data! "keyword:user" (user-added/imported)
    Dim c As New Collection
    Dim base As Collection, k As Variant
    Dim roleUse As String
    Dim arr As Variant
    Dim i As Long, item As String

    ' Start with hardcoded defaults
    Set base = DefaultKeywordBank()
    For Each k In base
        SafeAddToCollection c, CStr(k)
    Next k

    ' Merge Data sheet defaults
    arr = GetKeywordList("keyword:default")
    For i = LBound(arr) To UBound(arr)
        item = Trim$(CStr(arr(i)))
        If Len(item) > 0 Then SafeAddToCollection c, item
    Next i

    ' Determine role to use
    roleUse = Trim$(roleContext)
    If Len(roleUse) = 0 Then roleUse = Trim$(GetActiveRole())
    If Len(roleUse) > 0 Then
        arr = GetKeywordList("keyword:" & roleUse)
        For i = LBound(arr) To UBound(arr)
            item = Trim$(CStr(arr(i)))
            If Len(item) > 0 Then SafeAddToCollection c, item
        Next i
        ' Also add heuristic variations
        SafeAddToCollection c, roleUse
        SafeAddToCollection c, "experience in " & roleUse
        SafeAddToCollection c, roleUse & " skills"
    End If

    ' Merge user-added list
    arr = GetKeywordList("keyword:user")
    For i = LBound(arr) To UBound(arr)
        item = Trim$(CStr(arr(i)))
        If Len(item) > 0 Then SafeAddToCollection c, item
    Next i

    Set BuildKeywordBank = c
End Function

Private Sub SafeAddToCollection(ByRef c As Collection, ByVal value As String)
    ' Adds a unique lowercased value to collection (case-insensitive uniqueness)
    Dim v As Variant
    Dim normalized As String
    normalized = LCase$(Trim$(value))
    If Len(normalized) = 0 Then Exit Sub
    Dim exists As Boolean: exists = False
    For Each v In c
        If LCase$(CStr(v)) = normalized Then
            exists = True
            Exit For
        End If
    Next v
    If Not exists Then c.Add value
End Sub

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
