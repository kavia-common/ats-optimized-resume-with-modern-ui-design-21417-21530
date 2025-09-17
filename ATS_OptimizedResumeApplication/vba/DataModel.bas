Attribute VB_Name = "DataModel"
Option Explicit

' Module: DataModel
' Purpose:
'   Persist resume content, sections, and user preferences inside hidden sheets.
'   Provide CRUD methods with validation hooks and event notifications.

Private Const COL_KEY As Long = 1
Private Const COL_VALUE As Long = 2

' Keys for high-level data
Private Const KEY_ACTIVE_VERSION As String = "active_version"
Private Const KEY_ACTIVE_BRANCH As String = "active_branch"
Private Const KEY_SECTION_ORDER As String = "section_order"
Private Const KEY_USER_PREFS As String = "user_prefs"

' Keys for keyword profiles and role selection
Private Const KEY_ACTIVE_ROLE As String = "active_role"
' Convention:
'  - "keyword:default" -> comma-separated (or newline) list of default keywords (optional, complements built-ins)
'  - "keyword:{role}"   -> list of keywords for the given role/industry (comma-separated or newline)
'  - "keyword:user"     -> user-imported/added keywords (comma-separated or newline)

' PUBLIC_INTERFACE
Public Function GetActiveRole() As String
    ' Returns the active role/industry context; empty means none selected
    GetActiveRole = GetValue(KEY_ACTIVE_ROLE)
End Function

' PUBLIC_INTERFACE
Public Sub SetActiveRole(ByVal roleName As String)
    ' Sets the active role/industry context
    SetValue KEY_ACTIVE_ROLE, Trim$(roleName)
End Sub

Private Function NormalizeKeywordListRaw(ByVal raw As String) As String
    ' Normalizes list separators (comma/newline/semicolon) into newline-delimited storage
    Dim s As String
    s = Trim$(CStr(raw))
    If Len(s) = 0 Then
        NormalizeKeywordListRaw = ""
        Exit Function
    End If
    ' Replace semicolons with commas
    s = Replace(s, ";", ",")
    ' Replace CRLF and CR with LF for consistent splitting
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    ' Replace commas with LF to unify
    s = Replace(s, ",", vbLf)
    ' Collapse multiple LFs
    Do While InStr(1, s, vbLf & vbLf) > 0
        s = Replace(s, vbLf & vbLf, vbLf)
    Loop
    NormalizeKeywordListRaw = s
End Function

Private Function SerializeKeywordList(ByVal arr As Variant) As String
    ' Serializes array/collection to newline-delimited string
    Dim i As Long, s As String
    On Error GoTo done
    For i = LBound(arr) To UBound(arr)
        If Len(Trim$(CStr(arr(i)))) > 0 Then
            If Len(s) > 0 Then s = s & vbCrLf
            s = s & Trim$(CStr(arr(i)))
        End If
    Next i
done:
    SerializeKeywordList = s
End Function

' PUBLIC_INTERFACE
Public Function GetKeywordList(ByVal key As String) As Variant
    ' Loads and returns a 0-based array of keywords for the given key (e.g., "keyword:default", "keyword:Data Analyst")
    Dim raw As String, norm As String, parts() As String
    raw = GetValue(key)
    norm = NormalizeKeywordListRaw(raw)
    If Len(norm) = 0 Then
        GetKeywordList = Array()
        Exit Function
    End If
    parts = Split(norm, vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(parts(i))
    Next i
    GetKeywordList = parts
End Function

' PUBLIC_INTERFACE
Public Sub SetKeywordList(ByVal key As String, ByVal csvOrMultiline As String)
    ' Saves keywords for the given key; accepts comma-separated or newline separated input
    Dim norm As String
    norm = NormalizeKeywordListRaw(csvOrMultiline)
    SetValue key, norm
End Sub

' PUBLIC_INTERFACE
Public Function ListAvailableRoles() As Collection
    ' Scans Data sheet keys to list all roles with "keyword:{role}" entries
    Dim ws As Worksheet, lastRow As Long, i As Long, k As String, c As New Collection
    Set ws = SafeSheet(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        k = CStr(ws.Cells(i, COL_KEY).Value)
        If LCase$(Left$(k, 8)) = "keyword:" Then
            If LCase$(k) <> "keyword:default" And LCase$(k) <> "keyword:user" Then
                ' Role key - strip prefix
                c.Add Mid$(k, 9)
            End If
        End If
    Next i
    Set ListAvailableRoles = c
End Function

' PUBLIC_INTERFACE
Public Sub InitializeStorage()
    ' Initializes the Data sheet with headers if missing
    Dim ws As Worksheet
    Set ws = SafeSheet(SHEET_DATA)
    If ws.Cells(1, 1).Value <> "Key" Then
        ws.Cells.Clear
        ws.Cells(1, 1).Value = "Key"
        ws.Cells(1, 2).Value = "Value"
        ws.Columns(1).ColumnWidth = 40
        ws.Columns(2).ColumnWidth = 120
    End If
    EnsureDefaults
End Sub

Private Sub EnsureDefaults()
    If GetValue(KEY_ACTIVE_BRANCH) = "" Then
        SetValue KEY_ACTIVE_BRANCH, "main"
    End If
    If GetValue(KEY_ACTIVE_VERSION) = "" Then
        SetValue KEY_ACTIVE_VERSION, "v-0"
    End If
    If GetValue(KEY_SECTION_ORDER) = "" Then
        SetValue KEY_SECTION_ORDER, "Summary,Experience,Education,Skills,Projects,Certifications"
    End If
End Sub

' PUBLIC_INTERFACE
Public Function GetValue(key As String) As String
    ' Reads a single value from key-value store
    Dim ws As Worksheet, m As Variant, lastRow As Long, i As Long
    Set ws = SafeSheet(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i, COL_KEY).Value = key Then
            GetValue = ws.Cells(i, COL_VALUE).Value
            Exit Function
        End If
    Next i
    GetValue = ""
End Function

' PUBLIC_INTERFACE
Public Sub SetValue(key As String, value As String)
    ' Writes a single value into key-value store
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = SafeSheet(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1
    For i = 2 To lastRow
        If ws.Cells(i, COL_KEY).Value = key Then
            ws.Cells(i, COL_VALUE).Value = value
            Exit Sub
        End If
    Next i
    ws.Cells(lastRow + 1, COL_KEY).Value = key
    ws.Cells(lastRow + 1, COL_VALUE).Value = value
End Sub

' PUBLIC_INTERFACE
Public Function GetActiveBranch() As String
    ' Returns the current active branch
    GetActiveBranch = GetValue(KEY_ACTIVE_BRANCH)
End Function

' PUBLIC_INTERFACE
Public Sub SetActiveBranch(branchName As String)
    ' Sets the active branch
    SetValue KEY_ACTIVE_BRANCH, branchName
End Sub

' PUBLIC_INTERFACE
Public Function GetActiveVersionId() As String
    ' Returns the current active version id
    GetActiveVersionId = GetValue(KEY_ACTIVE_VERSION)
End Function

' PUBLIC_INTERFACE
Public Sub SetActiveVersionId(versionId As String)
    ' Sets the current active version id
    SetValue KEY_ACTIVE_VERSION, versionId
End Sub

' PUBLIC_INTERFACE
Public Function GetSectionOrder() As Collection
    ' Returns the current section order as a Collection
    Dim orderStr As String, parts() As String, i As Long
    Set GetSectionOrder = New Collection
    orderStr = GetValue(KEY_SECTION_ORDER)
    If Len(orderStr) = 0 Then Exit Function
    parts = Split(orderStr, ",")
    For i = LBound(parts) To UBound(parts)
        GetSectionOrder.Add Trim$(parts(i))
    Next i
End Function

' PUBLIC_INTERFACE
Public Sub SetSectionOrder(orderList As Collection)
    ' Persists a new section order
    Dim i As Long, s As String
    For i = 1 To orderList.Count
        s = s & orderList(i)
        If i < orderList.Count Then s = s & ","
    Next i
    SetValue KEY_SECTION_ORDER, s
End Sub

' PUBLIC_INTERFACE
Public Function LoadSectionData(sectionKey As String, versionId As String) As String
    ' Loads serialized JSON-like data for a section and version
    ' Storage key convention: section:{versionId}:{sectionKey}
    LoadSectionData = GetValue("section:" & versionId & ":" & sectionKey)
End Function

' PUBLIC_INTERFACE
Public Sub SaveSectionData(sectionKey As String, versionId As String, serialized As String)
    ' Saves serialized section data for a given version
    SetValue "section:" & versionId & ":" & sectionKey, serialized
End Sub

' PUBLIC_INTERFACE
Public Function ListCustomSections(versionId As String) As Collection
    ' Returns names of custom sections created by user
    Dim ws As Worksheet, lastRow As Long, i As Long, k As String
    Set ListCustomSections = New Collection
    Set ws = SafeSheet(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For i = 2 To lastRow
        k = CStr(ws.Cells(i, COL_KEY).Value)
        If InStr(1, k, "section:" & versionId & ":Custom:", vbTextCompare) = 1 Then
            ListCustomSections.Add Mid$(k, Len("section:" & versionId & ":") + 1)
        End If
    Next i
End Function

' PUBLIC_INTERFACE
Public Function SerializeArray(items() As String) As String
    ' Serializes a 0-based string array into a pipe-delimited string, escaping pipes
    Dim i As Long, s As String
    On Error GoTo EmptyArr
    For i = LBound(items) To UBound(items)
        s = s & Replace(items(i), "|", "||")
        If i < UBound(items) Then s = s & " | "
    Next i
    SerializeArray = s
    Exit Function
EmptyArr:
    SerializeArray = ""
End Function

' PUBLIC_INTERFACE
Public Function DeserializeArray(serialized As String) As Variant
    ' Deserializes a pipe-delimited string back to a 0-based array
    Dim parts() As String, i As Long
    If Len(serialized) = 0 Then
        DeserializeArray = Array()
        Exit Function
    End If
    parts = Split(serialized, " | ")
    For i = LBound(parts) To UBound(parts)
        parts(i) = Replace(parts(i), "||", "|")
    Next i
    DeserializeArray = parts
End Function
