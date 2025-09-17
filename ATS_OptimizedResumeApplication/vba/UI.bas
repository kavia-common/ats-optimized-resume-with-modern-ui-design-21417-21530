Attribute VB_Name = "UI"
Option Explicit

' Module: UI
' Purpose:
'   Orchestrates the in-Excel UI behavior using a dedicated UI sheet with controls.
'   Provides handlers for navigation, data binding, validation, and optimization displays.
'
' Cell address conventions on UI sheet (can be wired during setup):
'   B2: Active Template name
'   D2: Active Branch
'   F2: Active Version
'   A4: Section selector (data validation dropdown)
'   A6:A30: Section content editor (single-column; one cell per line)
'   H6: Keyword Score
'   H7:H20: Missing Keywords list
'   H3: Contextual Help
'   B3: Active Role (Data Validation dropdown)
'   G1: Manage Keywords label/button (assign UI.UI_ManageKeywords)

Private Const UI_CELL_TEMPLATE As String = "B2"
Private Const UI_CELL_BRANCH As String = "D2"
Private Const UI_CELL_VERSION As String = "F2"
Private Const UI_CELL_SECTION As String = "A4"
Private Const UI_RANGE_EDITOR As String = "A6:A30"
Private Const UI_CELL_KSCORE As String = "H6"
Private Const UI_RANGE_MISSING As String = "H7:H20"
Private Const UI_CELL_HELP As String = "H3"
Private Const UI_CELL_ROLE As String = "B3"

' Helper: join a collection into string
Private Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long, s As String
    For i = 1 To c.Count
        s = s & c(i)
        If i < c.Count Then s = s & sep
    Next i
    JoinCollection = s
End Function

' PUBLIC_INTERFACE
Public Sub SetupUI()
    ' Creates/initializes the UI worksheet with headings and data validations,
    ' including Role selector and Manage Keywords callout.
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

    ws.Range("A3").Value = "Role:"
    ws.Range(UI_CELL_ROLE).Value = GetActiveRole()

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
        If Len(s) > 0 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=s
            .IgnoreBlank = True
            .InCellDropdown = True
        End If
    End With

    ' Role dropdown from Data sheet roles (if any)
    Dim roles As Collection, rList As String, r As Long
    Set roles = ListAvailableRoles()
    rList = ""
    For r = 1 To roles.Count
        rList = rList & roles(r) & ","
    Next r
    If Len(rList) > 0 Then rList = Left$(rList, Len(rList) - 1)
    With ws.Range(UI_CELL_ROLE).Validation
        .Delete
        If Len(rList) > 0 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=rList
            .IgnoreBlank = True
            .InCellDropdown = True
        End If
    End With

    ' Set editor area style
    ws.Range(UI_RANGE_EDITOR).Style = STYLE_BODY
    ws.Range(UI_RANGE_EDITOR).EntireRow.RowHeight = 16

    ' Manage keywords hint/label
    ws.Range("G1").Value = "Manage Keywords ▶"
    ws.Range("G1").Style = STYLE_SUBHEADING

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

    ' Respect active role from UI (persist selection)
    Dim uiRole As String: uiRole = Trim$(CStr(ws.Range(UI_CELL_ROLE).Value))
    If uiRole <> GetActiveRole() Then
        SetActiveRole uiRole
    End If

    Dim bank As Collection
    Set bank = BuildKeywordBank(uiRole)
    Dim score As Double: score = ScoreKeywords(resumeAll, bank)
    ws.Range(UI_CELL_KSCORE).Value = "Keyword Score: " & CStr(score) & "/100"

    Dim missing As Collection: Set missing = KeywordGaps(resumeAll, bank)
    ws.Range(UI_RANGE_MISSING).ClearContents
    Dim i As Long, row As Long: row = ws.Range(UI_RANGE_MISSING).Row
    For i = 1 To WorksheetFunction.Min(missing.Count, ws.Range(UI_RANGE_MISSING).Rows.Count)
        ws.Cells(row + i - 1, ws.Range(UI_RANGE_MISSING).Column).Value = "• " & missing(i)
    Next i

    ' formatting violations for current section (show in help)
    Dim sec As String: sec = CStr(ws.Range(UI_CELL_SECTION).Value)
    Dim v As Collection: Set v = EnforceFormattingRules(sec, s)
    If v.Count > 0 Then
        ws.Range(UI_CELL_HELP).Value = "Issues:" & vbCrLf & JoinCollection(v, vbCrLf)
    Else
        ws.Range(UI_CELL_HELP).Value = "No formatting issues detected."
    End If
End Sub

' PUBLIC_INTERFACE
Public Sub UI_RoleChanged()
    ' Call this when the Role cell (B3) is changed; it persists role and refreshes insights
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim roleName As String: roleName = Trim$(CStr(ws.Range(UI_CELL_ROLE).Value))
    SetActiveRole roleName
    AccessibleAnnounce "Active role set to: " & IIf(Len(roleName) = 0, "(none)", roleName)
    UpdateKeywordInsights
End Sub

' PUBLIC_INTERFACE
Public Sub UI_ManageKeywords()
    ' Simple manager to select/edit:
    '  - Active Role
    '  - Role-specific keyword list
    '  - Default keyword list
    '  - User-added keyword list
    ' Accepts CSV-like or newline-separated input via InputBox prompts.
    Dim choice As String
    choice = InputBox( _
        "Keyword Manager:" & vbCrLf & _
        "1 = Select Active Role" & vbCrLf & _
        "2 = Edit Role Keywords (keyword:{active role})" & vbCrLf & _
        "3 = Edit Default Keywords (keyword:default)" & vbCrLf & _
        "4 = Edit User Keywords (keyword:user)" & vbCrLf & _
        "Enter 1-4:", "Manage Keywords")
    If StrPtr(choice) = 0 Then Exit Sub
    Select Case Trim$(choice)
        Case "1"
            Dim roles As Collection, i As Long, s As String, newRole As String
            Set roles = ListAvailableRoles()
            s = "Available roles:" & vbCrLf
            For i = 1 To roles.Count
                s = s & "• " & roles(i) & vbCrLf
            Next i
            newRole = InputBox(s & vbCrLf & "Enter role name to set (or new to create):", "Select Role", GetActiveRole())
            If StrPtr(newRole) = 0 Then Exit Sub
            SetActiveRole Trim$(newRole)
            SafeSheet(SHEET_UI).Range(UI_CELL_ROLE).Value = Trim$(newRole)
            AccessibleAnnounce "Role changed."
        Case "2"
            Dim roleName As String, existing As Variant, joined As String, inputList As String
            roleName = Trim$(GetActiveRole())
            If Len(roleName) = 0 Then
                roleName = InputBox("Enter role name to edit keywords for:", "Role Keywords")
                If StrPtr(roleName) = 0 Then Exit Sub
                roleName = Trim$(roleName)
                If Len(roleName) = 0 Then Exit Sub
                SetActiveRole roleName
                SafeSheet(SHEET_UI).Range(UI_CELL_ROLE).Value = roleName
            End If
            existing = GetKeywordList("keyword:" & roleName)
            joined = SerializeKeywordList(existing)
            inputList = InputBox("Enter CSV or newline-separated keywords for role '" & roleName & "':", "Edit Role Keywords", joined)
            If StrPtr(inputList) = 0 Then Exit Sub
            SetKeywordList "keyword:" & roleName, inputList
            AccessibleAnnounce "Saved role keywords for " & roleName
        Case "3"
            Dim exDef As Variant, jDef As String, inDef As String
            exDef = GetKeywordList("keyword:default")
            jDef = SerializeKeywordList(exDef)
            inDef = InputBox("Edit default keywords (affects all roles):", "Edit Default Keywords", jDef)
            If StrPtr(inDef) = 0 Then Exit Sub
            SetKeywordList "keyword:default", inDef
            AccessibleAnnounce "Saved default keywords."
        Case "4"
            Dim exUsr As Variant, jUsr As String, inUsr As String
            exUsr = GetKeywordList("keyword:user")
            jUsr = SerializeKeywordList(exUsr)
            inUsr = InputBox("Edit user-added keywords:", "Edit User Keywords", jUsr)
            If StrPtr(inUsr) = 0 Then Exit Sub
            SetKeywordList "keyword:user", inUsr
            AccessibleAnnounce "Saved user keywords."
        Case Else
            AccessibleAnnounce "Keyword Manager canceled."
    End Select
    UpdateKeywordInsights
End Sub
