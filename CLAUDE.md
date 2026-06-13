# PowerPoint VBA → Dataverse MCP client

## Goal
A VBA client (runs inside PowerPoint, but logic is host-agnostic) that signs in to
Microsoft Entra ID and calls the **Dataverse remote MCP server** at `/api/mcp` over
Streamable HTTP / JSON-RPC 2.0. Use case: drive Dataverse data/schema operations from
PowerPoint macros (e.g. query data and render it onto slides).

## Status: VALIDATED end-to-end ✅
The full chain was proven live from PowerShell + `curl.exe` against the target
environment (no PowerPoint needed for validation):
- Device code request → succeeds (confirms client/tenant + public client flows enabled)
- Browser sign-in → token granted with scope `…/api/mcp/mcp.tools`
- `initialize` → HTTP 200, server = "Microsoft Dataverse MCP Server v1.0.0", session id issued
  (confirms the app's client id IS on the environment allow-list)
- `notifications/initialized` → HTTP 202
- `tools/list` → **15 tools** returned
The VBA modules implement exactly this sequence; importing into PowerPoint should "just work".

## Environment / config (already baked into the code)
- **MCP URL:** `https://orge51b1ce8.crm22.dynamics.com/api/mcp`  (in `McpClient.bas`)
- **Tenant ID:** `ecd4deab-a1de-4d9d-a02a-2c87dfa5ecf2`  (in `McpAuth.bas`)
- **Client ID:** `bbe33aa1-2f52-42a4-9e44-17cad8f2392d`  (Entra app reg; in `McpAuth.bas`)
- **Scope:** `https://orge51b1ce8.crm22.dynamics.com/api/mcp/mcp.tools offline_access openid profile`
- **Protocol version:** `2025-06-18`
- These IDs are not secrets (public client). The only secret is the refresh token,
  which is stored DPAPI-encrypted, never in this repo.
- **These are the author's values.** `Presentation1.pptm` ships with the modules already
  imported, so it carries its OWN embedded copies of these constants. To target a different
  environment you must (1) edit `TENANT_ID`/`CLIENT_ID`/`MCP_SCOPE` in `McpAuth.bas` and
  `MCP_URL` in `McpClient.bas`, then (2) **re-import** the changed `.bas` into the deck
  (editing the file on disk does NOT update the copy inside the `.pptm`). Full walkthrough:
  README → *Point it at your own environment*.

## Auth design: OAuth 2.0 Device Code flow
Chosen because it's the simplest flow that works in VBA — no localhost redirect listener
(needed by auth-code+PKCE), no PKCE crypto, and it's a delegated user sign-in (client
credentials would be the wrong identity). `offline_access` yields a refresh token.

- Authority: `https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/{devicecode|token}`
- Flow: POST devicecode → show user_code + verification_uri → user signs in in browser →
  poll token endpoint (handle `authorization_pending` / `slow_down`).

### Refresh-token persistence (DPAPI)
- The rotating refresh token is encrypted with the Windows Data Protection API
  (`CryptProtectData`/`CryptUnprotectData`, per Windows user) and stored at
  `%APPDATA%\DataverseMcpVba\refresh_<clientid8>.dat`.
- Saved on every token response (Entra rotates refresh tokens). On startup it's loaded
  and redeemed silently; browser prompt only when missing/expired/revoked.
- A revoked token auto-deletes the file. `SignOut` clears cache + deletes file.
- Requires VBA7 (Office 2010+); older Office degrades to one sign-in per session.
- NOTE: DPAPI is per Windows-user + machine. On a **new computer** the saved token does
  NOT transfer — the first run there will require an interactive sign-in. That's expected.

## Files
- `McpAuth.bas`   — device code flow, token cache, refresh, DPAPI persistence.
- `McpClient.bas` — MCP handshake (`initialize` → `notifications/initialized`), `tools/list`,
  `tools/call`. Handles JSON and SSE-framed (`text/event-stream`) responses; captures and
  reuses the `Mcp-Session-Id` header; raises JSON-RPC errors.
- `modDemo.bas`   — entry points: `Demo_ListTools`, `Demo_CallTool`, `Demo_SilentReconnect`.
- `modDvmcp.bas`  — DVMCP text-box commands (see section below). The ONLY PowerPoint-specific
  module; everything else is host-agnostic.
- `README.md`     — user-facing setup.
- **`JsonConverter.bas`** — NOT in repo. Download from https://github.com/VBA-tools/VBA-JSON
  (VBA has no native JSON parser). Required at runtime; also needs the Microsoft Scripting
  Runtime reference.

## One-time Azure / Power Platform setup (ALREADY DONE for this env — verify if it breaks)
1. Entra app registration:
   - Authentication → Advanced → **Allow public client flows = Yes** (required for device code).
   - API permissions → Dynamics CRM → **mcp.tools** (delegated), consent granted.
   - No redirect URI needed for device code.
2. **Power Platform Admin Center** (separate from Azure, easy to miss):
   - Environment → Settings → Product → Features → *Dataverse Model Context Protocol* →
     **Advanced Settings** → add a client with the Application (client) ID, **Is Enabled = Yes**.
   - Without this, `/api/mcp` rejects an otherwise-valid token (401/403).

## Run it in PowerPoint (smoke test)
1. Download `JsonConverter.bas` (see Files).
2. PowerPoint → Save As → **.pptm** (macro-enabled).
3. Alt+F11 (VBE) → File → Import File… → import all 5 `.bas` (incl. JsonConverter).
4. Tools → References → tick **Microsoft Scripting Runtime**.
5. Run `Demo_ListTools` (F5) → browser opens, enter code, sign in → expect msgbox "15 tools".
6. Run `Demo_CallTool` → calls `describe('tables/')`, lists tables (read-only). No re-prompt.
7. Restart PowerPoint, run `Demo_SilentReconnect` → reconnects with no browser (DPAPI worked).
   - `SignOut` (Immediate window) then `Demo_ListTools` → prompt returns.

## The 15 MCP tools (from tools/list)
read_query, create_table, update_table, delete_table, create_record, update_record,
delete_record, search, upsert_skill, create_skill_resource, delete_skill, describe,
init_file_upload, commit_file_upload, file_download.

Verified input schemas (use for `tools/call`):
- `search`     → `{ "query": "<keywords>", "limit"?: int, "scope"?: string }`
- `describe`   → `{ "path": "tables/" | "tables/<name>" | "scopes/" | "skills/" | ... }`
- `read_query` → `{ "querytext": "<SELECT ...>" }`  (restricted SQL: basic SELECT, TOP,
  WHERE, ORDER BY, GROUP BY w/ COUNT/SUM/AVG/MIN/MAX, JOINs; NO subqueries, DISTINCT,
  HAVING, UNION, CAST, CONVERT, CASE, OFFSET, date functions.)

## DVMCP text-box commands (`modDvmcp.bas`)
Embed live Dataverse data in a slide: author a command as the **entire text** of a text box,
run a macro, and the box text is replaced by a rendered result. The original command is kept
on the shape so it can be re-run or restored. `read_query` is the focus/tested tool.

- **Syntax:** `DVMCP.tool(arg)`, where `arg` is either a quoted string mapped to the tool's
  primary parameter, or a raw `{JSON}` object passed straight through (escape hatch for any
  tool / multi-arg call). Example:
  `DVMCP.read_query("select firstname, lastname from contact order by createdon")`.
  - Primary-param mapping (simple-string form): `read_query→querytext`, `search→query`,
    `describe→path`. Other tools require the `{JSON}` form.
  - Wrap the arg in double quotes so SQL single-quote literals (`where x = 'a'`) pass through;
    a literal `"` is authored doubled (`""`); `'...'` is accepted as the alternate wrapper.
- **State lives in `Shape.Tags`** (PowerPoint has NO scalar `.Tag` — it's a name/value
  collection: `sh.Tags.Add`, `sh.Tags(name)`, `sh.Tags.Delete`). Tags:
  `DVMCPCommand` (original command; non-empty ⇒ managed), `DVMCPState` (`command`|`data`),
  `DVMCPError`, `DVMCPVersion`. Tags are the source of truth — never the visible text.
- **Two states:** *command* (editable command shown) ↔ *data* (result shown). Refresh on a
  data shape re-runs from the `DVMCPCommand` tag.
- **Macros:** `DVMCP_RefreshSelection|Slide|All` and `DVMCP_ShowCommands|Slide|All`.
  Refresh calls `McpInitialize` once, then loops matching shapes; a per-shape failure writes
  `DVMCP ERROR: …` into the box but **preserves `DVMCPCommand`**. ShowCommands restores the
  command text for editing.
- **Rendering:** results are written as a record list — field values joined with `, `,
  column headers omitted (box font set to Consolas). A `•` bullet prefixes each record only
  when there is more than one; a single record is shown as a plain line. GUID
  `…id` primary-key columns are hidden (`HIDE_ID_COLUMNS`, via `ShouldHideColumn`/`IsGuid`).
  The renderer is isolated in `RenderResult` / `RenderRowsAsList` — swap these to change the
  layout (e.g. a native PowerPoint table) without touching `ProcessShape`. The list path
  activates only if the tool's content parses as row objects (array, or
  `{rows|value|data|…:[…]}`); otherwise the raw text is shown verbatim (guaranteed fallback).
  - NOTE: `read_query`'s exact JSON shape was not captured at build time — if grids don't
    appear, inspect a real response (`Debug.Print`) and tune `AsRowCollection`.

## Validation harness (re-run without PowerPoint)
The whole flow can be smoke-tested from PowerShell + curl.exe. Pattern:
1. POST `…/oauth2/v2.0/devicecode` with `client_id` + `scope` → get `user_code`/`device_code`.
2. Sign in at the verification_uri, poll `…/oauth2/v2.0/token` with
   `grant_type=urn:ietf:params:oauth:grant-type:device_code`.
3. POST `initialize` to MCP_URL with headers `Authorization: Bearer`,
   `Accept: application/json, text/event-stream`, `MCP-Protocol-Version: 2025-06-18`;
   capture `Mcp-Session-Id` response header.
4. POST `notifications/initialized` (with session header) → expect 202.
5. POST `tools/list` / `tools/call` (with session header).
Note: use `curl.exe` not `Invoke-WebRequest` (the latter prompted on auth challenge in
non-interactive PowerShell). Strip SSE framing: keep lines starting with `data:`.

## Tip: develop/debug outside PowerPoint
`McpAuth` + `McpClient` have no PowerPoint dependency. Prototype in Excel VBA (nicer loop)
or a `.vbs` via `cscript.exe` (swap MsgBox/Debug.Print for WScript.Echo), then drop the
modules into PowerPoint unchanged.

## Possible next steps (not yet done)
- PowerPoint-specific helper: run `read_query` and render results onto a slide as a table.
- Surface tool input schemas in `Demo_ListTools` for discoverability.
- Fallback to auth-code + PKCE if a tenant Conditional Access policy blocks device code
  (symptom: AADSTS70011 / CA error on sign-in).
