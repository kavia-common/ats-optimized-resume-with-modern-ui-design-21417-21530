Attribute VB_Name = "Globals"
Option Explicit

' Module: Globals
' Purpose:
'   Centralized global constants, types, and helper functions used throughout the VBA project.
'   Avoid storing mutable global state here; prefer controlled property accessors.
' Notes:
'   - All PUBLIC_INTERFACE functions include docstrings per project requirements.
'   - Configuration should not be hard-coded; however, Excel VBA is self-contained. User paths are resolved dynamically.

' ===========================
' Global Constants and Types
' ===========================
Public Const APP_NAME As String = "ATS Optimized Resume Builder"
Public Const APP_VERSION As String = "1.0.0"
Public Const DEFAULT_TEMPLATE_NAME As String = "ModernMinimal"
Public Const SHEET_DATA As String = "Data"
Public Const SHEET_UI As String = "UI"
Public Const SHEET_VERSION_CTRL As String = "VersionHistory"
Public Const SHEET_DOCS As String = "HelpDocs"
Public Const SHEET_TEMPLATES As String = "Templates"
Public Const AUTO_SAVE_INTERVAL_MIN As Long = 5
Public Const MAX_UNDO_STACK As Long = 50

' standardized styles
Public Const STYLE_HEADING As String = "ATS_Heading"
Public Const STYLE_SUBHEADING As String = "ATS_Subheading"
Public Const STYLE_BODY As String = "ATS_Body"
Public Const STYLE_BULLET As String = "ATS_Bullet"

' Section keys
Public Enum ResumeSection
    rsSummary = 1
    rsExperience = 2
    rsEducation = 3
    rsSkills = 4
    rsProjects = 5
    rsCertifications = 6
    rsCustom = 7
End Enum

' Resume item types
Public Type WorkExperienceItem
    Title As String
    Company As String
    Location As String
    StartDate As Date
    EndDate As Date
    IsCurrent As Boolean
    Bullets() As String
    Keywords() As String
End Type

Public Type EducationItem
    Degree As String
    Institution As String
    Location As String
    GradDate As Date
    GPA As String
    Details() As String
End Type

Public Type ProjectItem
    Name As String
    Role As String
    StartDate As Date
    EndDate As Date
    URL As String
    Bullets() As String
    Skills() As String
End Type

Public Type ResumeVersionMeta
    VersionId As String
    Branch As String
    Title As String
    CreatedAt As Date
    CreatedBy As String
    Description As String
    ParentVersionId As String
End Type

' ===========================
' Utility helpers
' ===========================

' PUBLIC_INTERFACE
Public Function AppInfo() As String
    ' Provides basic app identity
    AppInfo = APP_NAME & " v" & APP_VERSION
End Function

' PUBLIC_INTERFACE
Public Function SafeSheet(name As String) As Worksheet
    ' Returns a worksheet by name or creates it if missing (hidden where appropriate)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = name
        If name = SHEET_DATA Or name = SHEET_VERSION_CTRL Or name = SHEET_DOCS Or name = SHEET_TEMPLATES Then
            ws.Visible = xlSheetVeryHidden
        End If
    End If
    Set SafeSheet = ws
End Function

' PUBLIC_INTERFACE
Public Sub EnsureNamedStyles()
    ' Creates/updates named styles for consistent formatting; ATS-friendly, accessible
    Dim st As Style
    On Error Resume Next
    Set st = ThisWorkbook.Styles(STYLE_HEADING)
    On Error GoTo 0
    If st Is Nothing Then
        Set st = ThisWorkbook.Styles.Add(STYLE_HEADING)
    End If
    With st
        .IncludeNumber = False
        .IncludeFont = True
        .IncludeAlignment = True
        .IncludeBorder = False
        .IncludePatterns = True
        .IncludeProtection = False
        .Font.name = "Calibri"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With

    On Error Resume Next
    Set st = ThisWorkbook.Styles(STYLE_SUBHEADING)
    On Error GoTo 0
    If st Is Nothing Then
        Set st = ThisWorkbook.Styles.Add(STYLE_SUBHEADING)
    End If
    With st
        .IncludeNumber = False
        .IncludeFont = True
        .IncludeAlignment = True
        .Font.name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(50, 50, 50)
    End With

    On Error Resume Next
    Set st = ThisWorkbook.Styles(STYLE_BODY)
    On Error GoTo 0
    If st Is Nothing Then
        Set st = ThisWorkbook.Styles.Add(STYLE_BODY)
    End If
    With st
        .IncludeNumber = False
        .IncludeFont = True
        .IncludeAlignment = True
        .Font.name = "Calibri"
        .Font.Size = 11
        .Font.Bold = False
        .Font.Color = RGB(40, 40, 40)
    End With

    On Error Resume Next
    Set st = ThisWorkbook.Styles(STYLE_BULLET)
    On Error GoTo 0
    If st Is Nothing Then
        Set st = ThisWorkbook.Styles.Add(STYLE_BULLET)
    End If
    With st
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .Font.name = "Calibri"
        .Font.Size = 11
        .Font.Color = RGB(40, 40, 40)
        ' Excel doesn't have bullets like Word; we prefix cells with "• "
    End With
End Sub

' PUBLIC_INTERFACE
Public Function NowIso() As String
    ' Returns ISO-like timestamp string for versioning metadata
    NowIso = Format(Now(), "yyyy-mm-dd\Thh:nn:ss")
End Function

' PUBLIC_INTERFACE
Public Function NewGuid() As String
    ' Returns a random identifier for versions; not a true GUID but sufficient in Excel
    Randomize
    NewGuid = Replace(NowIso & "-" & CStr(Int(Rnd() * 1000000)), ":", "")
End Function

' PUBLIC_INTERFACE
Public Sub AccessibleAnnounce(msg As String)
    ' Sends a message to the status bar for screen reader discovery
    Application.StatusBar = msg
End Sub

' PUBLIC_INTERFACE
Public Sub ResetStatusBar()
    ' Clears status bar
    Application.StatusBar = False
End Sub
