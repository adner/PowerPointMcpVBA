Attribute VB_Name = "modDvmcp"
Option Explicit

' =============================================================
'  modDvmcp - "DVMCP" text-box commands for PowerPoint
'  -----------------------------------------------------------
'  Author a command as the entire text of a text box, e.g.
'      DVMCP.read_query("select firstname, lastname from contact")
'  then run DVMCP_RefreshSelection (or _Slide / _All). The macro
'  calls the matching Dataverse MCP tool and REPLACES the box text
'  with a rendered (monospace) result. The original command is kept
'  on the shape via the PowerPoint Shape.Tags collection so it can be
'  re-run or restored for editing (DVMCP_ShowCommands*).
'
'  Two visible states, driven entirely by the tags (never by the
'  visible text): "command" (editable command) <-> "data" (result).
'
'  PowerPoint-specific code lives ONLY in this module. It reuses the
'  host-agnostic McpClient (McpInitialize / McpCallTool) + JsonConverter.
'  Depends on: McpClient, McpAuth, JsonConverter, Microsoft Scripting Runtime.
' =============================================================

' --- Scope of an operation ---
Private Enum DvScope
    dvSelection = 0
    dvSlide = 1
    dvAll = 2
End Enum

' read_query always returns the entity's primary-key column (e.g. "contactid",
' a GUID) even when not selected. Hide such columns from the grid to reduce
' clutter. Flip to False to show them.
Private Const HIDE_ID_COLUMNS As Boolean = True

' --- Tag names (the source of truth for a managed shape's state) ---
Private Const TAG_CMD As String = "DVMCPCommand"
Private Const TAG_STATE As String = "DVMCPState"
Private Const TAG_ERROR As String = "DVMCPError"
Private Const TAG_VERSION As String = "DVMCPVersion"

' ============================ Entry points ============================

Public Sub DVMCP_RefreshSelection()
    RefreshScope dvSelection
End Sub

Public Sub DVMCP_RefreshSlide()
    RefreshScope dvSlide
End Sub

Public Sub DVMCP_RefreshAll()
    RefreshScope dvAll
End Sub

Public Sub DVMCP_ShowCommands()
    ShowCommandsScope dvSelection
End Sub

Public Sub DVMCP_ShowCommandsSlide()
    ShowCommandsScope dvSlide
End Sub

Public Sub DVMCP_ShowCommandsAll()
    ShowCommandsScope dvAll
End Sub

' ============================ Drivers ============================

Private Sub RefreshScope(scope As DvScope)
    Dim shapes As Collection
    Set shapes = CollectShapes(scope)
    If shapes.Count = 0 Then
        Debug.Print "DVMCP: no command shapes found in " & ScopeName(scope) & "."
        Exit Sub
    End If

    ' Handshake once per run. Auth failure (e.g. the user cancels the
    ' device-code sign-in) aborts the whole run without touching any shape.
    On Error GoTo authFail
    McpInitialize
    On Error GoTo 0

    Dim sh As Shape, okCount As Long, errCount As Long, errLog As String
    Dim errNum As Long, errDesc As String
    For Each sh In shapes
        On Error Resume Next
        Err.Clear
        ProcessShape sh
        errNum = Err.Number: errDesc = Err.Description
        On Error GoTo 0
        If errNum <> 0 Then
            errCount = errCount + 1
            WriteError sh, errDesc
            If errCount <= 5 Then errLog = errLog & vbCrLf & "- " & errDesc
            Debug.Print "DVMCP shape error: " & errDesc
        Else
            okCount = okCount + 1
        End If
    Next sh

    Debug.Print "DVMCP: refreshed " & okCount & " of " & shapes.Count & _
                " shape(s), " & errCount & " error(s)." & errLog
    Exit Sub

authFail:
    ' Per-shape errors already surface in their text boxes; auth failure is logged.
    Debug.Print "DVMCP sign-in / handshake failed: " & Err.Description
End Sub

Private Sub ShowCommandsScope(scope As DvScope)
    Dim shapes As Collection
    Set shapes = CollectShapes(scope)
    Dim sh As Shape, n As Long
    For Each sh In shapes
        If Len(GetTag(sh, TAG_CMD)) > 0 Then
            RestoreCommand sh
            n = n + 1
        End If
    Next sh
    Debug.Print "DVMCP: restored " & n & " command(s) in " & ScopeName(scope) & "."
End Sub

' ============================ Per-shape ============================

Private Sub ProcessShape(sh As Shape)
    ' Command source depends on the current state:
    '   - "data": the box shows results, so re-run from the saved command tag.
    '   - otherwise (first run, or "command" view after ShowCommands): the visible
    '     text IS the command - read it so edits are picked up.
    Dim cmd As String
    If GetTag(sh, TAG_STATE) = "data" And Len(GetTag(sh, TAG_CMD)) > 0 Then
        cmd = GetTag(sh, TAG_CMD)
    Else
        If Not sh.HasTextFrame Then _
            Err.Raise vbObjectError + 27, "modDvmcp", "Shape has no text frame"
        cmd = TrimWs(sh.TextFrame.TextRange.Text)
    End If

    Dim toolName As String, argsJson As String
    ParseCommand cmd, toolName, argsJson          ' raises on a malformed command

    ' Persist the command now that it parsed cleanly - BEFORE the destructive
    ' write - so a later failure can never lose the user's input.
    SetTag sh, TAG_CMD, cmd
    SetTag sh, TAG_VERSION, "1"

    Dim resp As Object
    Set resp = McpCallTool(toolName, argsJson)

    WriteData sh, RenderResult(resp)
End Sub

Private Sub RestoreCommand(sh As Shape)
    Dim cmd As String
    cmd = GetTag(sh, TAG_CMD)
    If Len(cmd) = 0 Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    sh.TextFrame.TextRange.Text = cmd
    SetTag sh, TAG_STATE, "command"
End Sub

Private Sub WriteData(sh As Shape, text As String)
    sh.TextFrame.TextRange.Text = text
    On Error Resume Next
    sh.TextFrame.TextRange.Font.Name = "Consolas"   ' monospace so grids align
    On Error GoTo 0
    SetTag sh, TAG_STATE, "data"
    DeleteTag sh, TAG_ERROR                          ' clear any stale error
End Sub

Private Sub WriteError(sh As Shape, msg As String)
    SetTag sh, TAG_ERROR, msg
    ' Only overwrite the box if the command is safely saved in a tag; otherwise
    ' (e.g. a first-run parse failure) leave the author's text in place to fix.
    If Len(GetTag(sh, TAG_CMD)) > 0 Then
        If sh.HasTextFrame Then
            sh.TextFrame.TextRange.Text = "DVMCP ERROR: " & msg
            SetTag sh, TAG_STATE, "data"
        End If
    End If
End Sub

' ============================ Scope / shapes ============================

Private Function CollectShapes(scope As DvScope) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim sh As Shape, sld As Slide

    Select Case scope
        Case dvSelection
            Dim sel As Selection
            On Error Resume Next
            Set sel = ActiveWindow.Selection
            On Error GoTo 0
            If sel Is Nothing Then GoTo done
            If sel.Type <> ppSelectionShapes And sel.Type <> ppSelectionText Then GoTo done
            For Each sh In sel.ShapeRange
                If IsManaged(sh) Then result.Add sh
            Next sh

        Case dvSlide
            On Error Resume Next
            Set sld = ActiveWindow.View.Slide
            On Error GoTo 0
            If sld Is Nothing Then GoTo done
            For Each sh In sld.Shapes
                If IsManaged(sh) Then result.Add sh
            Next sh

        Case dvAll
            For Each sld In ActivePresentation.Slides
                For Each sh In sld.Shapes
                    If IsManaged(sh) Then result.Add sh
                Next sh
            Next sld
    End Select

done:
    Set CollectShapes = result
End Function

Private Function IsManaged(sh As Shape) As Boolean
    If Len(GetTag(sh, TAG_CMD)) > 0 Then
        IsManaged = True
        Exit Function
    End If
    If Not sh.HasTextFrame Then Exit Function
    If Not sh.TextFrame.HasText Then Exit Function
    If LCase$(Left$(TrimWs(sh.TextFrame.TextRange.Text), 6)) = "dvmcp." Then
        IsManaged = True
    End If
End Function

Private Function ScopeName(scope As DvScope) As String
    Select Case scope
        Case dvSelection: ScopeName = "the selection"
        Case dvSlide:     ScopeName = "the active slide"
        Case Else:        ScopeName = "the presentation"
    End Select
End Function

' ============================ Tag helpers ============================

Private Function GetTag(sh As Shape, tagName As String) As String
    GetTag = sh.Tags(tagName)                        ' returns "" when absent
End Function

Private Sub SetTag(sh As Shape, tagName As String, tagValue As String)
    DeleteTag sh, tagName                             ' delete-then-add = overwrite
    sh.Tags.Add tagName, tagValue
End Sub

Private Sub DeleteTag(sh As Shape, tagName As String)
    On Error Resume Next
    sh.Tags.Delete tagName
    On Error GoTo 0
End Sub

' ============================ Command parser ============================

' Parse "DVMCP.tool(args)" into a tool name and a JSON arguments string.
' args is either a single quoted "string" (mapped to the tool's primary
' parameter) or a raw {JSON} object (passed straight through).
Private Sub ParseCommand(ByVal raw As String, ByRef toolName As String, _
                         ByRef argsJson As String)
    Dim s As String
    s = TrimWs(raw)

    If LCase$(Left$(s, 6)) <> "dvmcp." Then
        Err.Raise vbObjectError + 20, "modDvmcp", _
                  "Not a DVMCP command (must start with 'DVMCP.')"
    End If

    Dim rest As String
    rest = Mid$(s, 7)

    Dim pOpen As Long, pClose As Long
    pOpen = InStr(rest, "(")
    If pOpen < 2 Then Err.Raise vbObjectError + 21, "modDvmcp", _
                                "Missing '(' after the tool name"
    pClose = InStrRev(rest, ")")
    If pClose <= pOpen Then Err.Raise vbObjectError + 22, "modDvmcp", _
                                      "Missing closing ')'"

    toolName = TrimWs(Left$(rest, pOpen - 1))
    Dim inner As String
    inner = TrimWs(Mid$(rest, pOpen + 1, pClose - pOpen - 1))

    ' Escape hatch: raw JSON object arguments (any tool, multi-arg).
    If Left$(inner, 1) = "{" Then
        Dim obj As Object
        If Not TryParseJson(inner, obj) Then
            Err.Raise vbObjectError + 23, "modDvmcp", _
                      "Argument looks like JSON but is not valid"
        End If
        argsJson = JsonConverter.ConvertToJson(obj)
        Exit Sub
    End If

    ' Friendly form: a single quoted string mapped to the tool's primary param.
    Dim q As String
    q = Left$(inner, 1)
    If q <> """" And q <> "'" Then
        Err.Raise vbObjectError + 24, "modDvmcp", _
                  "Argument must be a quoted string or a {JSON} object"
    End If
    If Len(inner) < 2 Or Right$(inner, 1) <> q Then
        Err.Raise vbObjectError + 25, "modDvmcp", "Unterminated quoted argument"
    End If

    Dim value As String
    value = Mid$(inner, 2, Len(inner) - 2)
    value = Replace(value, q & q, q)                 ' un-double embedded wrapper quotes

    Dim param As String
    param = PrimaryParamFor(toolName)
    If Len(param) = 0 Then
        Err.Raise vbObjectError + 26, "modDvmcp", _
                  "Tool '" & toolName & "' has no simple-syntax mapping; use {JSON} arguments"
    End If

    argsJson = "{""" & param & """:""" & JsonEscape(value) & """}"
End Sub

Private Function PrimaryParamFor(toolName As String) As String
    Select Case LCase$(toolName)
        Case "read_query": PrimaryParamFor = "querytext"
        Case "search":     PrimaryParamFor = "query"
        Case "describe":   PrimaryParamFor = "path"
        Case Else:         PrimaryParamFor = ""
    End Select
End Function

' Escape a raw string for embedding inside a JSON double-quoted literal.
Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr$(11), "\n")                   ' PowerPoint soft line break
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

' ============================ Result rendering ============================
' All formatting is isolated here so a different renderer (e.g. a native
' PowerPoint table) can be swapped in without touching ProcessShape.

Private Function RenderResult(resp As Object) As String
    Dim body As String
    body = ExtractText(resp)
    If Len(body) = 0 Then
        RenderResult = "(no content returned)"
        Exit Function
    End If

    ' If the content parses as tabular JSON, render a bulleted list of records;
    ' otherwise fall back to the raw text (markdown / plain / scalar).
    Dim parsed As Object
    If TryParseJson(body, parsed) Then
        Dim rows As Collection
        Set rows = AsRowCollection(parsed)
        If Not rows Is Nothing Then
            RenderResult = RenderRowsAsList(rows)
            Exit Function
        End If
    End If

    RenderResult = body
End Function

' Concatenate the text payloads from result.content (mirror Demo_CallTool).
Private Function ExtractText(resp As Object) As String
    Dim out As String
    Dim result As Object, content As Object, item As Variant
    On Error Resume Next
    Set result = resp("result")
    If Not result Is Nothing Then
        Set content = result("content")
        If Not content Is Nothing Then
            For Each item In content
                If item("type") = "text" Then out = out & item("text")
            Next item
        End If
    End If
    On Error GoTo 0
    ExtractText = out
End Function

' Coerce a parsed JSON value into a Collection of row Dictionaries, or Nothing.
Private Function AsRowCollection(parsed As Object) As Collection
    Dim rows As Collection
    If TypeOf parsed Is Collection Then
        Set rows = parsed
    ElseIf TypeOf parsed Is Scripting.Dictionary Then
        Dim k As Variant
        For Each k In Array("rows", "value", "data", "results", "records")
            If parsed.Exists(CStr(k)) Then
                If IsObject(parsed(CStr(k))) Then
                    If TypeOf parsed(CStr(k)) Is Collection Then
                        Set rows = parsed(CStr(k))
                        Exit For
                    End If
                End If
            End If
        Next k
    End If

    If rows Is Nothing Then Exit Function
    If rows.Count = 0 Then Exit Function
    If Not IsObject(rows(1)) Then Exit Function
    If Not (TypeOf rows(1) Is Scripting.Dictionary) Then Exit Function
    Set AsRowCollection = rows
End Function

' Render records one per line, field values joined and headers omitted. A
' bullet "• " is prefixed only when there is more than one record; a single
' record is shown as a plain line.
Private Function RenderRowsAsList(rows As Collection) As String
    Const MAX_ROWS As Long = 200
    Dim BULLET As String
    If rows.Count > 1 Then BULLET = ChrW$(8226) & " "     ' "• " (multi-record only)

    ' Column order = union of keys across the (capped) rows.
    Dim cols As Collection
    Set cols = New Collection
    Dim r As Long, rowObj As Object, k As Variant
    For r = 1 To rows.Count
        If r > MAX_ROWS Then Exit For
        Set rowObj = rows(r)
        For Each k In rowObj.Keys
            If Not CollHasKey(cols, CStr(k)) Then cols.Add CStr(k), CStr(k)
        Next k
    Next r
    If cols.Count = 0 Then
        RenderRowsAsList = "(0 columns)"
        Exit Function
    End If

    ' Drop auto-returned GUID "...id" primary-key columns to reduce clutter.
    Dim kept As Collection, col As Variant
    Set kept = New Collection
    For Each col In cols
        If Not ShouldHideColumn(CStr(col), rows) Then kept.Add CStr(col), CStr(col)
    Next col
    Set cols = kept
    If cols.Count = 0 Then
        RenderRowsAsList = "(no displayable columns)"
        Exit Function
    End If

    ' One bullet per record; field values joined with ", ".
    Dim sb As String, c As Long, parts As String
    For r = 1 To rows.Count
        If r > MAX_ROWS Then
            sb = sb & BULLET & "...(truncated, " & rows.Count & " rows total)" & vbCrLf
            Exit For
        End If
        Set rowObj = rows(r)
        parts = ""
        For c = 1 To cols.Count
            If c > 1 Then parts = parts & ", "
            parts = parts & CellText(rowObj, CStr(cols(c)))
        Next c
        sb = sb & BULLET & parts & vbCrLf
    Next r

    RenderRowsAsList = sb
End Function

' Hide a column only when its name ends with "id" AND it actually holds GUID
' values - so a primary key / GUID lookup is dropped, but a deliberately selected
' non-GUID column happening to end in "id" is preserved.
Private Function ShouldHideColumn(colName As String, rows As Collection) As Boolean
    If Not HIDE_ID_COLUMNS Then Exit Function
    If Right$(LCase$(colName), 2) <> "id" Then Exit Function
    Dim r As Long, v As String, sawValue As Boolean
    For r = 1 To rows.Count
        If r > 200 Then Exit For
        v = CellText(rows(r), colName)
        If Len(v) > 0 Then
            sawValue = True
            If Not IsGuid(v) Then Exit Function       ' non-GUID value -> keep column
        End If
    Next r
    ShouldHideColumn = sawValue                        ' hide only if GUID value(s) seen
End Function

Private Function IsGuid(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) <> 36 Then Exit Function
    Dim i As Long, ch As String
    For i = 1 To 36
        ch = Mid$(s, i, 1)
        If i = 9 Or i = 14 Or i = 19 Or i = 24 Then
            If ch <> "-" Then Exit Function
        ElseIf InStr("0123456789abcdefABCDEF", ch) = 0 Then
            Exit Function
        End If
    Next i
    IsGuid = True
End Function

Private Function CellText(rowObj As Object, key As String) As String
    If Not rowObj.Exists(key) Then Exit Function
    If IsObject(rowObj(key)) Then
        CellText = "[object]"
    Else
        Dim v As Variant
        v = rowObj(key)
        If IsNull(v) Or IsEmpty(v) Then
            CellText = ""
        Else
            CellText = CStr(v)
        End If
    End If
    CellText = Replace(Replace(CellText, vbCr, " "), vbLf, " ")
End Function

' ============================ Small utilities ============================

' True if a Collection already contains an item under the given string key.
Private Function CollHasKey(coll As Collection, key As String) As Boolean
    Dim tmp As Variant
    On Error Resume Next
    tmp = coll(key)
    CollHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function TryParseJson(ByVal s As String, ByRef outObj As Object) As Boolean
    On Error GoTo fail
    Set outObj = JsonConverter.ParseJson(s)
    TryParseJson = True
    Exit Function
fail:
    Set outObj = Nothing
    TryParseJson = False
End Function

' Trim spaces and control chars (CR/LF/TAB/VT) from both ends; keep internals.
Private Function TrimWs(ByVal s As String) As String
    Do While Len(s) > 0
        If Asc(Right$(s, 1)) > 32 Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    Do While Len(s) > 0
        If Asc(Left$(s, 1)) > 32 Then Exit Do
        s = Mid$(s, 2)
    Loop
    TrimWs = s
End Function
