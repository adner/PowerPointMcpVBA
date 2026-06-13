Attribute VB_Name = "McpClient"
Option Explicit

' =============================================================
'  McpClient - minimal Streamable HTTP MCP client (JSON-RPC 2.0)
'  Talks to the Dataverse remote MCP endpoint /api/mcp.
'  Depends on: McpAuth (bearer token) + JsonConverter (VBA-JSON).
' =============================================================

Public Const MCP_URL As String = "https://orge51b1ce8.crm22.dynamics.com/api/mcp"
Private Const PROTOCOL_VERSION As String = "2025-06-18"

Private mSessionId As String
Private mNextId As Long

' --- Handshake: must be called once before tools/list or tools/call ---
Public Sub McpInitialize()
    mNextId = 1
    mSessionId = ""
    Dim params As String
    params = "{""protocolVersion"":""" & PROTOCOL_VERSION & """," & _
             """capabilities"":{}," & _
             """clientInfo"":{""name"":""ppt-vba"",""version"":""0.1""}}"
    RpcCall "initialize", params            ' response carries Mcp-Session-Id header
    RpcNotify "notifications/initialized"   ' acknowledge per spec
End Sub

Public Function McpListTools() As Object
    Set McpListTools = RpcCall("tools/list", "")
End Function

' argsJson must be a JSON object literal, e.g. "{""tableName"":""account""}"
Public Function McpCallTool(toolName As String, Optional argsJson As String = "") As Object
    Dim params As String
    params = "{""name"":""" & toolName & """,""arguments"":" & _
             IIf(Len(argsJson) = 0, "{}", argsJson) & "}"
    Set McpCallTool = RpcCall("tools/call", params)
End Function

' ---------------- internals ----------------

Private Function RpcCall(method As String, paramsJson As String) As Object
    Dim id As Long: id = mNextId: mNextId = mNextId + 1
    Dim body As String
    body = "{""jsonrpc"":""2.0"",""id"":" & id & ",""method"":""" & method & """"
    If Len(paramsJson) > 0 Then body = body & ",""params"":" & paramsJson
    body = body & "}"
    Set RpcCall = McpPost(body)
End Function

Private Sub RpcNotify(method As String)
    McpPost "{""jsonrpc"":""2.0"",""method"":""" & method & """}"
End Sub

Private Function McpPost(jsonBody As String) As Object
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", MCP_URL, False
    http.setRequestHeader "Authorization", "Bearer " & McpAuth.GetAccessToken()
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json, text/event-stream"
    http.setRequestHeader "MCP-Protocol-Version", PROTOCOL_VERSION
    If Len(mSessionId) > 0 Then http.setRequestHeader "Mcp-Session-Id", mSessionId
    http.Send jsonBody

    Dim sid As String
    On Error Resume Next
    sid = http.getResponseHeader("Mcp-Session-Id")
    On Error GoTo 0
    If Len(sid) > 0 Then mSessionId = sid

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 2, "McpClient", _
                  "HTTP " & http.Status & ": " & http.responseText
    End If

    Dim payload As String
    If InStr(LCase$(http.getResponseHeader("Content-Type")), "text/event-stream") > 0 Then
        payload = ExtractSseJson(http.responseText)   ' SSE-framed response
    Else
        payload = http.responseText
    End If

    If Len(Trim$(payload)) = 0 Then
        Set McpPost = Nothing                          ' 202 Accepted for notifications
    Else
        Set McpPost = JsonConverter.ParseJson(payload)
        ' Surface JSON-RPC level errors (only present on failure responses)
        If McpPost.Exists("error") Then
            Dim e As Object
            Set e = McpPost("error")
            Err.Raise vbObjectError + 3, "McpClient", _
                      "JSON-RPC error " & e("code") & ": " & e("message")
        End If
    End If
End Function

' Pull the JSON out of the last SSE 'data:' event in the body.
Private Function ExtractSseJson(raw As String) As String
    Dim lines() As String, i As Long, dataBuf As String, lastJson As String, ln As String
    lines = Split(Replace(raw, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        If Left$(ln, 5) = "data:" Then
            dataBuf = dataBuf & Trim$(Mid$(ln, 6))
        ElseIf Len(Trim$(ln)) = 0 Then
            If Len(dataBuf) > 0 Then lastJson = dataBuf
            dataBuf = ""
        End If
    Next i
    If Len(dataBuf) > 0 Then lastJson = dataBuf
    ExtractSseJson = lastJson
End Function
