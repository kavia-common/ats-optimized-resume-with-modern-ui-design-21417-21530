Attribute VB_Name = "Accessibility"
Option Explicit

' Module: Accessibility
' Purpose:
'   Provide undo/redo stacks for UI editor, sheet protection utilities, and announce helpers.

Private Type StackItem
    Payload As String
End Type

Private UndoStack As Collection
Private RedoStack As Collection

Private Function GetEditorPayload() As String
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim rng As Range: Set rng = ws.Range("A6:A30")
    Dim s As String, c As Range
    For Each c In rng.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            If Len(s) > 0 Then s = s & vbCrLf
            s = s & CStr(c.Value)
        End If
    Next c
    GetEditorPayload = s
End Function

Private Sub SetEditorPayload(ByVal payload As String)
    Dim ws As Worksheet: Set ws = SafeSheet(SHEET_UI)
    Dim rng As Range: Set rng = ws.Range("A6:A30")
    rng.ClearContents
    Dim lines() As String, i As Long
    If Len(payload) = 0 Then Exit Sub
    lines = Split(payload, vbCrLf)
    For i = LBound(lines) To WorksheetFunction.Min(UBound(lines), rng.Rows.Count - 1)
        rng.Cells(i + 1, 1).Value = lines(i)
    Next i
End Sub

' PUBLIC_INTERFACE
Public Sub UndoPush()
    ' Push current editor payload onto undo stack
    If UndoStack Is Nothing Then Set UndoStack = New Collection
    Dim it As StackItem, s As String
    s = GetEditorPayload()
    it.Payload = s
    If UndoStack.Count >= MAX_UNDO_STACK Then
        ' Drop oldest by recreating collection
        Dim tmp As New Collection, i As Long
        For i = 2 To UndoStack.Count
            tmp.Add UndoStack(i)
        Next i
        Set UndoStack = tmp
    End If
    UndoStack.Add it
End Sub

' PUBLIC_INTERFACE
Public Sub Undo()
    ' Restores previous editor state and pushes current to redo stack
    If UndoStack Is Nothing Or UndoStack.Count = 0 Then
        AccessibleAnnounce "Nothing to undo."
        Exit Sub
    End If
    Dim current As String: current = GetEditorPayload()
    If RedoStack Is Nothing Then Set RedoStack = New Collection
    Dim redoIt As StackItem: redoIt.Payload = current
    RedoStack.Add redoIt

    Dim it As StackItem
    it = UndoStack(UndoStack.Count)
    SetEditorPayload it.Payload
    UndoStack.Remove UndoStack.Count
    AccessibleAnnounce "Undo performed."
End Sub

' PUBLIC_INTERFACE
Public Sub Redo()
    ' Restores next redo state
    If RedoStack Is Nothing Or RedoStack.Count = 0 Then
        AccessibleAnnounce "Nothing to redo."
        Exit Sub
    End If
    Dim it As StackItem
    it = RedoStack(RedoStack.Count)
    SetEditorPayload it.Payload
    RedoStack.Remove RedoStack.Count
    AccessibleAnnounce "Redo performed."
End Sub

' PUBLIC_INTERFACE
Public Sub ProtectDataSheets()
    ' Protect internal sheets to prevent accidental user edits
    On Error Resume Next
    SafeSheet(SHEET_DATA).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    SafeSheet(SHEET_VERSION_CTRL).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    SafeSheet(SHEET_DOCS).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    SafeSheet(SHEET_TEMPLATES).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0
End Sub

' PUBLIC_INTERFACE
Public Sub UnprotectDataSheets()
    ' Unprotect internal sheets (e.g., during maintenance)
    On Error Resume Next
    SafeSheet(SHEET_DATA).Unprotect
    SafeSheet(SHEET_VERSION_CTRL).Unprotect
    SafeSheet(SHEET_DOCS).Unprotect
    SafeSheet(SHEET_TEMPLATES).Unprotect
    On Error GoTo 0
End Sub
