Attribute VB_Name = "ExportPrint"
Option Explicit

' Module: ExportPrint
' Purpose:
'   Create exportable view of resume and export to PDF or send to printer.

' PUBLIC_INTERFACE
Public Function BuildExportSheet(Optional ByVal forPrint As Boolean = False) As Worksheet
    ' Assembles the resume content into a new/clean worksheet named "Export"
    Dim ws As Worksheet, i As Long, row As Long
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Export").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.name = "Export"
    row = 1

    EnsureNamedStyles
    Dim order As Collection: Set order = GetSectionOrder()
    Dim ver As String: ver = GetActiveVersionId()
    Dim sec As String, payload As String, lines() As String, ln As Variant

    For i = 1 To order.Count
        sec = order(i)
        payload = LoadSectionData(sec, ver)
        If Len(payload) = 0 Then GoTo NextSec
        ws.Cells(row, 1).Value = UCase$(sec)
        ws.Cells(row, 1).Style = STYLE_HEADING
        row = row + 1
        lines = Split(payload, vbCrLf)
        For Each ln In lines
            If Len(Trim$(CStr(ln))) = 0 Then GoTo skip
            If Left$(CStr(ln), 2) = "• " Or Left$(CStr(ln), 1) = "-" Then
                ws.Cells(row, 1).Value = CStr(ln)
                ws.Cells(row, 1).Style = STYLE_BULLET
            Else
                ws.Cells(row, 1).Value = CStr(ln)
                ws.Cells(row, 1).Style = STYLE_BODY
            End If
            row = row + 1
skip:
        Next ln
        row = row + 1
NextSec:
    Next i

    With ws.PageSetup
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    ws.Columns(1).ColumnWidth = 100 / 7 ' approx to fit letter width

    If forPrint Then
        ws.PageSetup.PrintArea = ws.UsedRange.Address
    End If
    Set BuildExportSheet = ws
End Function

' PUBLIC_INTERFACE
Public Sub ExportActiveVersionToPDF(ByVal outputPath As String)
    ' Exports the built export sheet to PDF at the given path
    Dim ws As Worksheet
    Set ws = BuildExportSheet(True)
    On Error GoTo errh
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outputPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    AccessibleAnnounce "Exported PDF: " & outputPath
    Exit Sub
errh:
    AccessibleAnnounce "PDF export failed. Check path permissions."
End Sub

' PUBLIC_INTERFACE
Public Function BuildPlainTextResume() As String
    ' Builds a plain text representation suitable for ATS parsing tests
    Dim order As Collection: Set order = GetSectionOrder()
    Dim ver As String: ver = GetActiveVersionId()
    Dim i As Long, sec As String, payload As String, s As String
    For i = 1 To order.Count
        sec = order(i)
        payload = LoadSectionData(sec, ver)
        If Len(payload) > 0 Then
            s = s & UCase$(sec) & vbCrLf & payload & vbCrLf & vbCrLf
        End If
    Next i
    BuildPlainTextResume = s
End Function
```

Explanation: UI orchestration for guided forms, navigation, real-time validation feedback, and contextual help. This module wires up button handlers (to be assigned to Form Controls on UI sheet) and provides section CRUD and keyword scoring display.

````write file="ats-optimized-resume-with-modern-ui-design-21417-21530/ATS_OptimizedResumeApplication/vba/UI.bas"
Attribute VB_Name = "UI"
Option Explicit

' Module: UI
' Purpose:
'   Orchestrates the in-Excel UI behavior using a dedicated UI sheet with controls.
'   Provides handlers for navigation, data binding, validation, and optimization displays.

' Cell address conventions on UI sheet (can be wired during setup):
'   B2: Active Template name
'   D2: Active Branch
'   F2: Active Version
'   A4: Section selector (data validation dropdown)
'   A6:A30: Section content editor (single-column; one cell per line)
'   H6: Keyword Score
'   H7:H20: Missing Keywords list
'   H3: Contextual Help

Private Const UI_CELL_TEMPLATE As String = "B2"
Private Const UI_CELL_BRANCH As String = "D2"
Private Const UI_CELL_VERSION As String = "F2"
Private Const UI_CELL_SECTION As String = "A4"
Private Const UI_RANGE_EDITOR As String = "A6:A30"
Private Const UI_CELL_KSCORE As String = "H6"
Private Const UI_RANGE_MISSING As String = "H7:H20"
Private Const UI_CELL_HELP As String = "H3"

' PUBLIC_INTERFACE
Public Sub SetupUI()
    ' Creates/initializes the UI worksheet with headings and data validations
    Dim ws As Worksheet
    Set ws = SafeSheet(SHEET_UI)
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ApplyTemplate ActiveTemplateName()

    ws.Range("B1").Value = "Template:"
    ws.Range(UI_CELL_TEMPLATE).Value = ActiveTemplateName()
    ws.Range("D1").Value = "Branch:"
    ws.Range(UI_CELL_BRANCH).Value = GetActiveBranch()
    ws.Range("F1").Value = "Version:"
    ws.Range(UI_CELL_VERSION).Value = GetActiveVersionId()

    ws.Range("A4").Value = "Summary"
    ws.Range("A5").Value = "Edit lines below (each line is a bullet if prefixed with '• '):"
    ws.Range("A5").Style = STYLE_SUBHEADING
    ws.Range("H3").Value = "Help will appear here."

    ' Build dropdown for section selector
    Dim order As Collection, i As Long, s As String
    Set order = GetSectionOrder()
    For i = 1 To order.Count
        s = s & order(i) & ","
    Next i
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    With ws.Range(UI_CELL_SECTION).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=s
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Set editor area style
    ws.Range(UI_RANGE_EDITOR).Style = STYLE_BODY
    ws.Range(UI_RANGE_EDITOR).EntireRow.RowHeight = 16

    AccessibleAnnounce "UI initialized. Select a section and edit content."
End Sub

' PUBLIC_INTERFACE
Public Sub UI_SectionChanged()
    ' Call when section dropdown changes to load section data into editor
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim sec As String: sec = CStr(ws.Range(UI_CELL_SECTION).Value)
    Dim ver As String: ver = GetActiveVersionId()
    Dim payload As String: payload = LoadSectionData(sec, ver)

    ' populate lines
    ws.Range(UI_RANGE_EDITOR).ClearContents
    If Len(payload) > 0 Then
        Dim lines() As String, i As Long
        lines = Split(payload, vbCrLf)
        For i = LBound(lines) To WorksheetFunction.Min(UBound(lines), ws.Range(UI_RANGE_EDITOR).Rows.Count - 1)
            ws.Range(UI_RANGE_EDITOR).Cells(i + 1, 1).Value = lines(i)
        Next i
    End If

    ' update help
    ws.Range(UI_CELL_HELP).Value = "Editing section: " & sec & vbCrLf & _
        "Tips: " & HelpTipForField(sec)
    UpdateKeywordInsights
End Sub

' PUBLIC_INTERFACE
Public Sub UI_SaveSection()
    ' Saves current editor content into active version
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim sec As String: sec = CStr(ws.Range(UI_CELL_SECTION).Value)
    If Len(sec) = 0 Then
        AccessibleAnnounce "Select a section first."
        Exit Sub
    End If
    Dim ver As String: ver = GetActiveVersionId()
    Dim rng As Range: Set rng = ws.Range(UI_RANGE_EDITOR)
    Dim s As String, c As Range
    For Each c In rng.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            If Len(s) > 0 Then s = s & vbCrLf
            s = s & CStr(c.Value)
        End If
    Next c
    SaveSectionData sec, ver, s
    AccessibleAnnounce "Saved section '" & sec & "'."
    UpdateKeywordInsights
End Sub

' PUBLIC_INTERFACE
Public Sub UI_AddBulletPrefix()
    ' Adds "• " prefix to selected editor cells for bullet standardization
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.Selection
    On Error GoTo 0
    If rng Is Nothing Then Set rng = ws.Range(UI_RANGE_EDITOR)
    Dim c As Range
    For Each c In rng.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            If Left$(CStr(c.Value), 2) <> "• " Then
                c.Value = "• " & c.Value
            End If
        End If
    Next c
    AccessibleAnnounce "Applied bullet prefix."
End Sub

' PUBLIC_INTERFACE
Public Sub UI_NewVersion()
    ' Creates a manual version with a description prompt
    Dim desc As String
    desc = InputBox("Enter version description:", "Save Version")
    If StrPtr(desc) = 0 Then Exit Sub ' canceled
    Dim id As String
    id = CreateVersion("Manual Save", desc, False)
    SafeSheet(SHEET_UI).Range(UI_CELL_VERSION).Value = id
End Sub

' PUBLIC_INTERFACE
Public Sub UI_ExportPDF()
    ' Prompts for path and exports PDF
    Dim f As Variant
    f = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.Path & "\resume.pdf", FileFilter:="PDF Files (*.pdf), *.pdf")
    If f = False Then Exit Sub
    ExportActiveVersionToPDF CStr(f)
End Sub

' PUBLIC_INTERFACE
Public Sub UI_CompareVersions()
    ' Prompts for two version IDs and shows a diff summary
    Dim a As String, b As String, s As String
    a = InputBox("Enter base version ID:", "Compare Versions")
    If StrPtr(a) = 0 Then Exit Sub
    b = InputBox("Enter compare version ID:", "Compare Versions")
    If StrPtr(b) = 0 Then Exit Sub
    s = CompareVersions(a, b)
    MsgBox s, vbInformation, "Diff Summary"
End Sub

' PUBLIC_INTERFACE
Public Sub UI_CreateBranch()
    ' Prompts for branch name and switches to it
    Dim br As String
    br = InputBox("Enter new branch name:", "Create Branch")
    If StrPtr(br) = 0 Then Exit Sub
    CreateBranch br
    SafeSheet(SHEET_UI).Range(UI_CELL_BRANCH).Value = br
End Sub

' PUBLIC_INTERFACE
Public Sub UI_MergeIntoActive()
    ' Prompts for source version to merge into active
    Dim src As String, tgt As String
    src = InputBox("Enter source version ID to merge from:", "Merge Versions")
    If StrPtr(src) = 0 Then Exit Sub
    tgt = GetActiveVersionId()
    MergeVersions src, tgt
    AccessibleAnnounce "Merged changes into active version."
End Sub

' PUBLIC_INTERFACE
Public Sub UI_ApplyTemplate()
    ' Applies template typed in template cell, or prompts if empty
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim name As String: name = CStr(ws.Range(UI_CELL_TEMPLATE).Value)
    If Len(Trim$(name)) = 0 Then
        name = InputBox("Template name:", "Apply Template", ActiveTemplateName())
        If StrPtr(name) = 0 Then Exit Sub
    End If
    ApplyTemplate name
    ws.Range(UI_CELL_TEMPLATE).Value = ActiveTemplateName()
End Sub

' PUBLIC_INTERFACE
Public Sub UpdateKeywordInsights()
    ' Calculates keyword score and missing keywords based on current section + overall text
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim rng As Range: Set rng = ws.Range(UI_RANGE_EDITOR)
    Dim s As String, c As Range
    For Each c In rng.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            If Len(s) > 0 Then s = s & vbCrLf
            s = s & CStr(c.Value)
        End If
    Next c

    Dim resumeAll As String
    resumeAll = BuildPlainTextResume()
    If Len(s) > 0 Then resumeAll = resumeAll & vbCrLf & s

    Dim bank As Collection: Set bank = BuildKeywordBank("")
    Dim score As Double: score = ScoreKeywords(resumeAll, bank)
    ws.Range(UI_CELL_KSCORE).Value = "Keyword Score: " & CStr(score) & "/100"

    Dim missing As Collection: Set missing = KeywordGaps(resumeAll, bank)
    ws.Range(UI_RANGE_MISSING).ClearContents
    Dim i As Long, row As Long: row = ws.Range(UI_RANGE_MISSING).Row
    For i = 1 To WorksheetFunction.Min(missing.Count, ws.Range(UI_RANGE_MISSING).Rows.Count)
        ws.Cells(row + i - 1, ws.Range(UI_RANGE_MISSING).Column).Value = "• " & missing(i)
    Next i

    ' formatting violations for current section
    Dim sec As String: sec = CStr(ws.Range(UI_CELL_SECTION).Value)
    Dim v As Collection: Set v = EnforceFormattingRules(sec, s)
    If v.Count > 0 Then
        ws.Range(UI_CELL_HELP).Value = "Issues:" & vbCrLf & JoinCollection(v, vbCrLf)
    Else
        ws.Range(UI_CELL_HELP).Value = "No formatting issues detected."
    End If
End Sub

Private Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long, s As String
    For i = 1 To c.Count
        s = s & c(i)
        If i < c.Count Then s = s & sep
    Next i
    JoinCollection = s
End Function
