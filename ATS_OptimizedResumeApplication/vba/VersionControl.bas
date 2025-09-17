Attribute VB_Name = "VersionControl"
Option Explicit

' Module: VersionControl
' Purpose:
'   Maintain version history with auto-save, manual save, branching, compare, and restore.
'   Stores serialized section data per version and a metadata log.

Private mTimerRunning As Boolean
Private mNextAutoSave As Date

Private Sub EnsureHistorySheet()
    Dim ws As Worksheet
    Set ws = SafeSheet(SHEET_VERSION_CTRL)
    If ws.Cells(1, 1).Value <> "VersionId" Then
        ws.Cells.Clear
        ws.Range("A1:H1").Value = Array("VersionId", "Branch", "Title", "Description", "Timestamp", "Author", "ParentVersionId", "Active")
        ws.Columns("A:H").ColumnWidth = 24
    End If
End Sub

' PUBLIC_INTERFACE
Public Sub StartAutoSave()
    ' Starts autosave timer with configured interval
    If mTimerRunning Then Exit Sub
    mTimerRunning = True
    mNextAutoSave = Now() + TimeSerial(0, AUTO_SAVE_INTERVAL_MIN, 0)
    Application.OnTime mNextAutoSave, "VersionControl_AutoSaveTick", , True
    AccessibleAnnounce "Auto-save enabled every " & AUTO_SAVE_INTERVAL_MIN & " minutes."
End Sub

' PUBLIC_INTERFACE
Public Sub StopAutoSave()
    ' Stops autosave timer
    On Error Resume Next
    If mTimerRunning Then
        Application.OnTime mNextAutoSave, "VersionControl_AutoSaveTick", , False
        mTimerRunning = False
        AccessibleAnnounce "Auto-save disabled."
    End If
    On Error GoTo 0
End Sub

Public Sub VersionControl_AutoSaveTick()
    ' Internal callback
    On Error Resume Next
    CreateVersion "Auto-save", "Automatic save at " & NowIso, True
    ' Reschedule
    If mTimerRunning Then
        mNextAutoSave = Now() + TimeSerial(0, AUTO_SAVE_INTERVAL_MIN, 0)
        Application.OnTime mNextAutoSave, "VersionControl_AutoSaveTick", , True
    End If
    On Error GoTo 0
End Sub

' PUBLIC_INTERFACE
Public Function CreateVersion(ByVal title As String, ByVal description As String, Optional silent As Boolean = False) As String
    ' Creates a new version, persists metadata, and copies section data from active version
    EnsureHistorySheet
    Dim newId As String, branch As String, parent As String
    newId = "v-" & NewGuid()
    branch = GetActiveBranch()
    parent = GetActiveVersionId()

    ' copy section payloads
    Dim order As Collection, i As Long, secKey As String, payload As String
    Set order = GetSectionOrder()
    For i = 1 To order.Count
        secKey = order(i)
        payload = LoadSectionData(secKey, parent)
        If Len(payload) > 0 Then
            SaveSectionData secKey, newId, payload
        End If
    Next i

    ' register metadata
    Dim ws As Worksheet, r As Long
    Set ws = SafeSheet(SHEET_VERSION_CTRL)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = newId
    ws.Cells(r, 2).Value = branch
    ws.Cells(r, 3).Value = title
    ws.Cells(r, 4).Value = description
    ws.Cells(r, 5).Value = NowIso
    ws.Cells(r, 6).Value = Environ$("USERNAME")
    ws.Cells(r, 7).Value = parent
    ws.Cells(r, 8).Value = IIf(Not silent, "Active", "")

    SetActiveVersionId newId
    If Not silent Then AccessibleAnnounce "Created version " & newId & " on branch " & branch
    CreateVersion = newId
End Function

' PUBLIC_INTERFACE
Public Sub RestoreVersion(ByVal versionId As String)
    ' Sets an existing version as active by ID
    If Len(versionId) = 0 Then Exit Sub
    SetActiveVersionId versionId
    AccessibleAnnounce "Restored version " & versionId
End Sub

' PUBLIC_INTERFACE
Public Function CompareVersions(ByVal lhs As String, ByVal rhs As String) As String
    ' Returns a simple diff summary by section of char-length differences
    Dim order As Collection, i As Long, sec As String, a As String, b As String, s As String
    Set order = GetSectionOrder()
    For i = 1 To order.Count
        sec = order(i)
        a = LoadSectionData(sec, lhs)
        b = LoadSectionData(sec, rhs)
        If a <> b Then
            s = s & sec & ": Δchars=" & Abs(Len(a) - Len(b)) & vbCrLf
        End If
    Next i
    If Len(s) = 0 Then s = "No differences."
    CompareVersions = s
End Function

' PUBLIC_INTERFACE
Public Sub CreateBranch(ByVal name As String)
    ' Switch to a new branch; active version remains
    If Len(Trim$(name)) = 0 Then
        AccessibleAnnounce "Branch name required."
        Exit Sub
    End If
    SetActiveBranch name
    AccessibleAnnounce "Switched to branch " & name
End Sub

' PUBLIC_INTERFACE
Public Function MergeVersions(ByVal sourceVersion As String, ByVal targetVersion As String) As String
    ' Performs a naive merge: if target section empty, copy from source; if both non-empty, append unique lines
    Dim order As Collection, i As Long, sec As String
    Dim a As String, b As String, merged As String
    Set order = GetSectionOrder()
    For i = 1 To order.Count
        sec = order(i)
        a = LoadSectionData(sec, sourceVersion)
        b = LoadSectionData(sec, targetVersion)
        If Len(b) = 0 Then
            merged = a
        ElseIf Len(a) = 0 Then
            merged = b
        Else
            merged = MergeTextLines(a, b)
        End If
        SaveSectionData sec, targetVersion, merged
    Next i
    AccessibleAnnounce "Merged " & sourceVersion & " -> " & targetVersion
    MergeVersions = targetVersion
End Function

Private Function MergeTextLines(ByVal a As String, ByVal b As String) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim line As Variant, arrA() As String, arrB() As String
    arrA = Split(a, vbCrLf)
    arrB = Split(b, vbCrLf)
    For Each line In arrA
        If Len(Trim$(CStr(line))) > 0 Then dict(Trim$(CStr(line))) = True
    Next line
    For Each line In arrB
        If Len(Trim$(CStr(line))) > 0 Then dict(Trim$(CStr(line))) = True
    Next line
    Dim k As Variant, s As String
    For Each k In dict.Keys
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & CStr(k)
    Next k
    MergeTextLines = s
End Function
