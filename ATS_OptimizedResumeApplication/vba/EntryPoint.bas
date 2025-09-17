Attribute VB_Name = "EntryPoint"
Option Explicit

' Module: EntryPoint
' Purpose:
'   Initializes the application on workbook open and exposes public routines to setup and run the UI.

' PUBLIC_INTERFACE
Public Sub App_Initialize()
    ' Initializes storage, templates, UI, autosave, and documentation
    InitializeStorage
    EnsureNamedStyles
    ApplyTemplate ActiveTemplateName()
    SetupUI
    ExportStyleGuide
    StartAutoSave
    ProtectDataSheets
    AccessibleAnnounce AppInfo() & " initialized."
End Sub

' PUBLIC_INTERFACE
Public Sub App_ShowHelp()
    ' Displays the in-app help overview message and points user to HelpDocs sheet
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_DOCS)
    ws.Visible = xlSheetVisible
    MsgBox "Help and style guide is available on the 'HelpDocs' sheet." & vbCrLf & _
           "FAQ, troubleshooting, and tips are included. Keyboard hints:" & vbCrLf & _
           " - Use Tab/Shift+Tab to navigate form cells" & vbCrLf & _
           " - Use Ctrl+Z / Ctrl+Y for Undo/Redo (or buttons if provided)", vbInformation, "Help"
End Sub

' PUBLIC_INTERFACE
Public Sub App_Shutdown()
    ' Clean-up tasks on closing the workbook
    StopAutoSave
    ResetStatusBar
End Sub
