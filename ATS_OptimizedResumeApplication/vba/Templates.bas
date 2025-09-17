Attribute VB_Name = "Templates"
Option Explicit

' Module: Templates
' Purpose:
'   Manage a library of ATS-friendly templates and enforce styling on the resume UI sheet.

Private Type TemplateDef
    Name As String
    HeadingColor As Long
    AccentColor As Long
    BodyColor As Long
    FontName As String
    HeadingSize As Integer
    BodySize As Integer
End Type

Private mTemplates As Collection
Private mActiveTemplate As String

Private Sub EnsureTemplates()
    If mTemplates Is Nothing Then
        Set mTemplates = New Collection
        Dim t As TemplateDef

        t.Name = "ModernMinimal"
        t.HeadingColor = RGB(0, 0, 0)
        t.AccentColor = RGB(30, 80, 160)
        t.BodyColor = RGB(40, 40, 40)
        t.FontName = "Calibri"
        t.HeadingSize = 16
        t.BodySize = 11
        mTemplates.Add t, t.Name

        t.Name = "CleanProfessional"
        t.HeadingColor = RGB(10, 10, 10)
        t.AccentColor = RGB(0, 120, 110)
        t.BodyColor = RGB(45, 45, 45)
        t.FontName = "Segoe UI"
        t.HeadingSize = 15
        t.BodySize = 11
        mTemplates.Add t, t.Name
    End If
    If Len(mActiveTemplate) = 0 Then mActiveTemplate = DEFAULT_TEMPLATE_NAME
End Sub

Private Function GetTemplate(ByVal name As String) As TemplateDef
    EnsureTemplates
    Dim td As TemplateDef
    On Error Resume Next
    td = mTemplates(name)
    On Error GoTo 0
    GetTemplate = td
End Function

' PUBLIC_INTERFACE
Public Function ListTemplates() As Collection
    ' Returns template names
    EnsureTemplates
    Dim c As New Collection, i As Long
    For i = 1 To mTemplates.Count
        Dim td As TemplateDef
        td = mTemplates(i)
        c.Add td.Name
    Next i
    Set ListTemplates = c
End Function

' PUBLIC_INTERFACE
Public Sub ApplyTemplate(ByVal name As String)
    ' Applies the template to UI sheet: styles and color scheme
    EnsureTemplates
    Dim td As TemplateDef
    td = GetTemplate(name)
    If Len(td.Name) = 0 Then
        AccessibleAnnounce "Template not found; using default."
        td = GetTemplate(DEFAULT_TEMPLATE_NAME)
    End If
    mActiveTemplate = td.Name
    EnsureNamedStyles

    ' Update named styles to reflect template palette
    With ThisWorkbook.Styles(STYLE_HEADING).Font
        .name = td.FontName
        .Size = td.HeadingSize
        .Bold = True
        .Color = td.HeadingColor
    End With
    With ThisWorkbook.Styles(STYLE_SUBHEADING).Font
        .name = td.FontName
        .Size = td.BodySize + 1
        .Bold = True
        .Color = td.AccentColor
    End With
    With ThisWorkbook.Styles(STYLE_BODY).Font
        .name = td.FontName
        .Size = td.BodySize
        .Color = td.BodyColor
    End With
    With ThisWorkbook.Styles(STYLE_BULLET).Font
        .name = td.FontName
        .Size = td.BodySize
        .Color = td.BodyColor
    End With

    ' Apply to UI sheet header areas if present
    Dim ws As Worksheet
    Set ws = SafeSheet(SHEET_UI)
    With ws
        .Cells.ClearFormats
        .Range("A1").Value = APP_NAME
        .Range("A1").Style = STYLE_HEADING
        .Range("A2").Value = "Template: " & mActiveTemplate
        .Range("A2").Style = STYLE_SUBHEADING
        .Columns("A:H").ColumnWidth = 30
        .Rows("1:3").RowHeight = 18
    End With
    AccessibleAnnounce "Applied template: " & mActiveTemplate
End Sub

' PUBLIC_INTERFACE
Public Function ActiveTemplateName() As String
    ' Returns current template name
    ActiveTemplateName = mActiveTemplate
End Function

' PUBLIC_INTERFACE
Public Sub ExportStyleGuide()
    ' Writes style guide and tips into hidden HelpDocs sheet for in-app docs
    Dim ws As Worksheet
    Set ws = SafeSheet(SHEET_DOCS)
    ws.Cells.Clear
    ws.Range("A1").Value = "In-App Style Guide"
    ws.Range("A1").Style = STYLE_HEADING

    ws.Range("A3").Value = "ATS Formatting Rules"
    ws.Range("A3").Style = STYLE_SUBHEADING
    ws.Range("A4").Value = "• Use plain text, no images or tables" & vbCrLf & _
                           "• Keep bullets <= 200 chars" & vbCrLf & _
                           "• Consistent headings and section order" & vbCrLf & _
                           "• Avoid columns that may confuse parsers"
    ws.Range("A4").Style = STYLE_BODY

    ws.Range("A8").Value = "Templates"
    ws.Range("A8").Style = STYLE_SUBHEADING
    Dim i As Long, c As Collection
    Set c = ListTemplates()
    Dim s As String: s = ""
    For i = 1 To c.Count
        s = s & "• " & c(i) & vbCrLf
    Next i
    ws.Range("A9").Value = s
    ws.Range("A9").Style = STYLE_BODY
End Sub
