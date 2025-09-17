Attribute VB_Name = "Validation"
Option Explicit

' Module: Validation
' Purpose:
'   Real-time validation for form fields and contextual help lookup.

' PUBLIC_INTERFACE
Public Function IsValidDateText(ByVal s As String) As Boolean
    ' Returns True if s can be parsed to a Date and is within reasonable range
    On Error GoTo bad
    If Len(Trim$(s)) = 0 Then GoTo bad
    Dim d As Date
    d = CDate(s)
    If Year(d) < 1950 Or Year(d) > 2100 Then GoTo bad
    IsValidDateText = True
    Exit Function
bad:
    IsValidDateText = False
End Function

' PUBLIC_INTERFACE
Public Function IsValidEmail(ByVal s As String) As Boolean
    ' Basic email regex approximation
    IsValidEmail = (InStr(1, s, "@") > 1 And InStrRev(s, ".") > InStr(1, s, "@") + 1)
End Function

' PUBLIC_INTERFACE
Public Function IsValidURL(ByVal s As String) As Boolean
    ' Simple URL validation; for tips only
    IsValidURL = (Left$(LCase$(s), 7) = "http://" Or Left$(LCase$(s), 8) = "https://")
End Function

' PUBLIC_INTERFACE
Public Function RequireNonEmpty(ByVal s As String) As Boolean
    ' True if string not empty
    RequireNonEmpty = (Len(Trim$(s)) > 0)
End Function

' PUBLIC_INTERFACE
Public Function HelpTipForField(ByVal fieldName As String) As String
    ' Returns contextual help text for field
    Select Case LCase$(fieldName)
        Case "title"
            HelpTipForField = "Your job title or role, e.g., 'Senior Analyst'. Keep it concise and aligned to target role."
        Case "company"
            HelpTipForField = "Organization name. Use official name; avoid abbreviations unless widely recognized."
        Case "startdate", "enddate"
            HelpTipForField = "Use a consistent date format, e.g., 'Jan 2022'."
        Case "bullets"
            HelpTipForField = "Start bullets with action verbs, quantify results, and keep lines <= 200 characters."
        Case "skills"
            HelpTipForField = "List specific, relevant skills; group by category if helpful."
        Case "summary"
            HelpTipForField = "2–4 lines summarizing your value proposition and target keywords."
        Case Else
            HelpTipForField = "Provide clear, concise information. Avoid images, tables, or complex formatting."
    End Select
End Function

' PUBLIC_INTERFACE
Public Function ValidateDateRange(ByVal startText As String, ByVal endText As String, ByVal isCurrent As Boolean) As String
    ' Returns empty string if valid; otherwise error message
    If Not IsValidDateText(startText) Then
        ValidateDateRange = "Invalid start date."
        Exit Function
    End If
    If isCurrent Then
        ValidateDateRange = ""
        Exit Function
    End If
    If Not IsValidDateText(endText) Then
        ValidateDateRange = "Invalid end date."
        Exit Function
    End If
    If CDate(endText) < CDate(startText) Then
        ValidateDateRange = "End date must be after start date."
    Else
        ValidateDateRange = ""
    End If
End Function
