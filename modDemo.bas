Attribute VB_Name = "modDemo"
Option Explicit

' Entry points to test from the VBE (F5) or Developer > Macros.

' 1) List the tools the Dataverse MCP server exposes. (Pure smoke test.)
Public Sub Demo_ListTools()
    McpInitialize
    Dim resp As Object, tools As Object, t As Variant, s As String, n As Long
    Set resp = McpListTools()
    Set tools = resp("result")("tools")
    For Each t In tools
        n = n + 1
        s = s & n & ". " & t("name") & vbCrLf
    Next t
    MsgBox "Connected. " & n & " Dataverse MCP tools:" & vbCrLf & vbCrLf & s, vbInformation
    Debug.Print s
End Sub

' 2) Call a read-only tool: describe('tables/') lists all tables in the environment.
'    Verified schema: describe takes { "path": "..." }.
Public Sub Demo_CallTool()
    McpInitialize
    Dim resp As Object
    Set resp = McpCallTool("describe", "{""path"":""tables/""}")

    Dim content As Object, item As Variant, out As String
    Set content = resp("result")("content")
    For Each item In content
        If item("type") = "text" Then out = out & item("text") & vbCrLf
    Next item

    If Len(out) > 1200 Then out = Left$(out, 1200) & vbCrLf & "...(truncated)"
    MsgBox out, vbInformation, "describe('tables/') result"
    Debug.Print out
End Sub

' 3) Prove persistence: forces use of the saved refresh token (no browser if cached).
'    Run Demo_ListTools once (sign in), restart PowerPoint, then run this.
Public Sub Demo_SilentReconnect()
    McpInitialize
    MsgBox "Reconnected without a browser prompt - refresh token persisted via DPAPI.", vbInformation
End Sub
