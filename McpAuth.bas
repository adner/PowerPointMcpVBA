Attribute VB_Name = "McpAuth"
Option Explicit

' =============================================================
'  McpAuth - OAuth 2.0 Device Code flow for Entra ID (Microsoft)
'  Acquires a delegated access token for the Dataverse MCP server.
'  Persists the rotating refresh token across sessions using DPAPI
'  (Windows Data Protection API) - encrypted per Windows user.
'  Depends on: JsonConverter (VBA-JSON) - see README.
' =============================================================

' ====== CONFIGURE THESE ======
Public Const TENANT_ID As String = "ecd4deab-a1de-4d9d-a02a-2c87dfa5ecf2"
Public Const CLIENT_ID As String = "bbe33aa1-2f52-42a4-9e44-17cad8f2392d"
Public Const MCP_SCOPE As String = _
    "https://orge51b1ce8.crm22.dynamics.com/api/mcp/mcp.tools offline_access openid profile"
' =============================

Private Const AUTHORITY As String = "https://login.microsoftonline.com/"

#If VBA7 Then
    Private Declare PtrSafe Sub SleepMs Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#Else
    Private Declare Sub SleepMs Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#End If

' --- DPAPI declarations (Type/Declare/Const must live in the declarations
'     section, before any procedure - VBA rule) ---
#If VBA7 Then
Private Type DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type

Private Declare PtrSafe Function CryptProtectData Lib "crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, ByVal szDataDescr As LongPtr, _
    ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, _
    ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB) As Long

Private Declare PtrSafe Function CryptUnprotectData Lib "crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, ByVal ppszDataDescr As LongPtr, _
    ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, _
    ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal length As Long)

Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

Private Const CRYPTPROTECT_UI_FORBIDDEN As Long = &H1
#End If

' In-memory cache
Private mAccessToken As String
Private mRefreshToken As String
Private mExpiresAt As Date

' Returns a valid bearer token. Order: cached -> refresh (mem or disk) -> interactive.
Public Function GetAccessToken() As String
    If Len(mAccessToken) > 0 And Now < mExpiresAt Then
        GetAccessToken = mAccessToken
        Exit Function
    End If

    If Len(mRefreshToken) = 0 Then mRefreshToken = LoadRefreshToken()   ' from DPAPI file

    If Len(mRefreshToken) > 0 Then
        If RedeemRefreshToken() Then
            GetAccessToken = mAccessToken
            Exit Function
        End If
    End If

    If DeviceCodeFlow() Then
        GetAccessToken = mAccessToken
    Else
        Err.Raise vbObjectError + 1, "McpAuth", "Failed to acquire access token."
    End If
End Function

' Clear cache and forget the persisted refresh token (forces fresh sign-in).
Public Sub SignOut()
    mAccessToken = "": mRefreshToken = "": mExpiresAt = 0
    DeleteRefreshToken
End Sub

Private Function DeviceCodeFlow() As Boolean
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", AUTHORITY & TENANT_ID & "/oauth2/v2.0/devicecode", False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send "client_id=" & CLIENT_ID & "&scope=" & UrlEncode(MCP_SCOPE)

    If http.Status <> 200 Then
        Debug.Print "devicecode error " & http.Status & ": " & http.responseText
        Exit Function
    End If

    Dim dc As Object
    Set dc = JsonConverter.ParseJson(http.responseText)

    Dim deviceCode As String, userCode As String, verifyUrl As String
    Dim interval As Long, expiresIn As Long, waited As Long
    deviceCode = dc("device_code")
    userCode = dc("user_code")
    verifyUrl = dc("verification_uri")
    interval = CLng(dc("interval"))
    expiresIn = CLng(dc("expires_in"))

    OpenUrl verifyUrl
    MsgBox "Sign in to Dataverse:" & vbCrLf & vbCrLf & _
           "1. A browser opened at:  " & verifyUrl & vbCrLf & _
           "2. Enter this code:  " & userCode & vbCrLf & vbCrLf & _
           "Complete sign-in, then click OK and wait a moment.", _
           vbInformation, "Dataverse MCP sign-in"

    Do While waited < expiresIn
        SleepMs interval * 1000
        waited = waited + interval

        http.Open "POST", AUTHORITY & TENANT_ID & "/oauth2/v2.0/token", False
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.Send "grant_type=urn:ietf:params:oauth:grant-type:device_code" & _
                  "&client_id=" & CLIENT_ID & "&device_code=" & deviceCode

        Dim res As Object
        Set res = JsonConverter.ParseJson(http.responseText)

        If http.Status = 200 Then
            StoreToken res
            DeviceCodeFlow = True
            Exit Function
        End If

        Select Case CStr(SafeGet(res, "error"))
            Case "authorization_pending"            ' keep polling
            Case "slow_down":            interval = interval + 5
            Case Else
                Debug.Print "token error: " & http.responseText
                Exit Function
        End Select
    Loop
End Function

Private Function RedeemRefreshToken() As Boolean
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", AUTHORITY & TENANT_ID & "/oauth2/v2.0/token", False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send "grant_type=refresh_token&client_id=" & CLIENT_ID & _
              "&refresh_token=" & UrlEncode(mRefreshToken) & _
              "&scope=" & UrlEncode(MCP_SCOPE)

    If http.Status = 200 Then
        StoreToken JsonConverter.ParseJson(http.responseText)
        RedeemRefreshToken = True
    Else
        Debug.Print "refresh failed " & http.Status & ": " & http.responseText
        mRefreshToken = ""
        DeleteRefreshToken                          ' stale/revoked - drop it
    End If
End Function

Private Sub StoreToken(res As Object)
    mAccessToken = res("access_token")
    mExpiresAt = DateAdd("s", CLng(res("expires_in")) - 60, Now)   ' 60s safety buffer
    Dim rt As Variant
    rt = SafeGet(res, "refresh_token")
    If Not IsEmpty(rt) Then
        mRefreshToken = CStr(rt)
        SaveRefreshToken mRefreshToken            ' refresh tokens rotate - persist each time
    End If
End Sub

Private Function SafeGet(o As Object, key As String) As Variant
    On Error Resume Next
    SafeGet = o(key)
    On Error GoTo 0
End Function

Public Function UrlEncode(s As String) As String
    Dim i As Long, c As String, code As Long, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or c = "-" Or c = "_" Or c = "." Or c = "~" Then
            out = out & c
        ElseIf c = " " Then
            out = out & "%20"
        Else
            out = out & "%" & Right$("0" & Hex$(code), 2)
        End If
    Next i
    UrlEncode = out
End Function

Private Sub OpenUrl(url As String)
    On Error Resume Next
    CreateObject("WScript.Shell").Run "rundll32 url.dll,FileProtocolHandler " & url, 1, False
    On Error GoTo 0
End Sub

' ===================== DPAPI refresh-token persistence =====================
' Encrypted with the current Windows user's key; only this user on this
' machine can decrypt. Stored under %APPDATA%\DataverseMcpVba\.

#If VBA7 Then

Private Function TokenFilePath() As String
    Dim folder As String
    folder = Environ$("APPDATA") & "\DataverseMcpVba"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    TokenFilePath = folder & "\refresh_" & Left$(CLIENT_ID, 8) & ".dat"
End Function

Private Sub SaveRefreshToken(t As String)
    On Error GoTo fail
    Dim plain() As Byte, enc() As Byte
    plain = StrConv(t, vbFromUnicode)              ' token is ASCII - 1 byte/char
    enc = DpapiProtect(plain)
    Dim f As Integer: f = FreeFile
    Open TokenFilePath() For Output As #f
    Print #f, ToBase64(enc)
    Close #f
    Exit Sub
fail:
    Debug.Print "SaveRefreshToken failed: " & Err.Description
End Sub

Private Function LoadRefreshToken() As String
    On Error GoTo fail
    Dim p As String: p = TokenFilePath()
    If Len(Dir$(p)) = 0 Then Exit Function
    Dim f As Integer, ln As String
    f = FreeFile
    Open p For Input As #f
    Line Input #f, ln
    Close #f
    If Len(ln) = 0 Then Exit Function
    Dim enc() As Byte, dec() As Byte
    enc = FromBase64(ln)
    dec = DpapiUnprotect(enc)
    LoadRefreshToken = StrConv(dec, vbUnicode)
    Exit Function
fail:
    Debug.Print "LoadRefreshToken failed: " & Err.Description
    LoadRefreshToken = ""
End Function

Private Sub DeleteRefreshToken()
    On Error Resume Next
    Dim p As String: p = TokenFilePath()
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0
End Sub

Private Function DpapiProtect(ByRef data() As Byte) As Byte()
    Dim inBlob As DATA_BLOB, outBlob As DATA_BLOB
    inBlob.cbData = UBound(data) - LBound(data) + 1
    inBlob.pbData = VarPtr(data(LBound(data)))
    If CryptProtectData(inBlob, 0, 0, 0, 0, CRYPTPROTECT_UI_FORBIDDEN, outBlob) = 0 Then
        Err.Raise vbObjectError + 10, "McpAuth", "CryptProtectData failed"
    End If
    Dim out() As Byte
    ReDim out(0 To outBlob.cbData - 1)
    CopyMemory VarPtr(out(0)), outBlob.pbData, outBlob.cbData
    LocalFree outBlob.pbData
    DpapiProtect = out
End Function

Private Function DpapiUnprotect(ByRef data() As Byte) As Byte()
    Dim inBlob As DATA_BLOB, outBlob As DATA_BLOB
    inBlob.cbData = UBound(data) - LBound(data) + 1
    inBlob.pbData = VarPtr(data(LBound(data)))
    If CryptUnprotectData(inBlob, 0, 0, 0, 0, CRYPTPROTECT_UI_FORBIDDEN, outBlob) = 0 Then
        Err.Raise vbObjectError + 11, "McpAuth", "CryptUnprotectData failed"
    End If
    Dim out() As Byte
    ReDim out(0 To outBlob.cbData - 1)
    CopyMemory VarPtr(out(0)), outBlob.pbData, outBlob.cbData
    LocalFree outBlob.pbData
    DpapiUnprotect = out
End Function

Private Function ToBase64(ByRef b() As Byte) As String
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = b
    ToBase64 = node.Text
End Function

Private Function FromBase64(s As String) As Byte()
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.Text = s
    FromBase64 = node.nodeTypedValue
End Function

#Else  ' Pre-VBA7 (Office 2007 and earlier): no persistence, per-session sign-in.

Private Function LoadRefreshToken() As String
    LoadRefreshToken = ""
End Function

Private Sub SaveRefreshToken(t As String)
End Sub

Private Sub DeleteRefreshToken()
End Sub

#End If
