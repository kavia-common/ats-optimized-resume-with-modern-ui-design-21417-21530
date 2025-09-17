Attribute VB_Name = "Documentation"
Option Explicit

' Module: Documentation
' Purpose:
'   Generate and update comprehensive user guidance: manual, FAQ, troubleshooting within HelpDocs sheet.

' PUBLIC_INTERFACE
Public Sub BuildDocumentation()
    ' Rebuilds the HelpDocs sheet content with manual, FAQ, troubleshooting
    Dim ws As Worksheet
    Set ws = SafeSheet(SHEET_DOCS)
    ws.Cells.Clear

    ws.Range("A1").Value = "ATS Optimized Resume Builder - User Manual"
    ws.Range("A1").Style = ThisWorkbook.Styles("ATS_Heading")

    ws.Range("A3").Value = "Getting Started"
    ws.Range("A3").Style = ThisWorkbook.Styles("ATS_Subheading")
    ws.Range("A4").Value = "1. Go to UI sheet" & vbCrLf & _
                           "2. Select a section from dropdown" & vbCrLf & _
                           "3. Enter content line-by-line (prefix '• ' for bullets)" & vbCrLf & _
                           "4. Click Save Section" & vbCrLf & _
                           "5. Check Keyword Score and fix missing keywords."

    ws.Range("A8").Value = "ATS Optimization Tips"
    ws.Range("A8").Style = ThisWorkbook.Styles("ATS_Subheading")
    ws.Range("A9").Value = "• Use relevant keywords naturally" & vbCrLf & _
                           "• Avoid tables, images, and complex formatting" & vbCrLf & _
                           "• Keep bullets concise and results-driven"

    ws.Range("A13").Value = "FAQ"
    ws.Range("A13").Style = ThisWorkbook.Styles("ATS_Subheading")
    ws.Range("A14").Value = "Q: How do I export to PDF?" & vbCrLf & _
                            "A: Click 'Export PDF' on the UI and choose a location." & vbCrLf & vbCrLf & _
                            "Q: How do I manage versions?" & vbCrLf & _
                            "A: Use 'Save Version', 'Compare Versions', and 'Restore Version' features." & vbCrLf & vbCrLf & _
                            "Q: How do I change the template?" & vbCrLf & _
                            "A: Update the Template cell on UI or use 'Apply Template'."

    ws.Range("A22").Value = "Troubleshooting"
    ws.Range("A22").Style = ThisWorkbook.Styles("ATS_Subheading")
    ws.Range("A23").Value = "Issue: Low keyword score" & vbCrLf & _
                            "Fix: Review missing keywords and include them where relevant." & vbCrLf & vbCrLf & _
                            "Issue: PDF export fails" & vbCrLf & _
                            "Fix: Ensure write permissions to folder and try again." & vbCrLf & vbCrLf & _
                            "Issue: Validation errors" & vbCrLf & _
                            "Fix: Check date formats and required fields."

    ws.Columns("A:A").ColumnWidth = 100 / 7
End Sub
