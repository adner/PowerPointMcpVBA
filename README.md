# PowerPoint VBA → Dataverse MCP client

A VBA client that runs **inside PowerPoint**, signs in to Microsoft Entra ID with the
OAuth 2.0 **device code** flow, and calls the **Dataverse remote MCP server** (`/api/mcp`)
over Streamable HTTP / JSON-RPC 2.0. Use it to query live Dataverse data from a macro and
render it straight onto a slide.

> ⚠️ **This repo is pre-configured for one specific Dataverse environment** (the author's).
> The tenant ID, client ID, MCP URL and scope are **hard-coded** in the VBA source. To use it
> against *your* environment you must edit those constants and re-import the modules — see
> **[Point it at your own environment](#point-it-at-your-own-environment)** below.

---

## What's in the repo

| File | Purpose |
|------|---------|
| `Presentation1.pptm` | Macro-enabled PowerPoint deck with **all modules already imported** (incl. `JsonConverter`). The fastest way to try the demo. |
| `McpAuth.bas` | Device-code sign-in, token cache, refresh, DPAPI persistence. **Holds `TENANT_ID` / `CLIENT_ID` / `MCP_SCOPE`.** |
| `McpClient.bas` | MCP handshake, `tools/list`, `tools/call`. **Holds `MCP_URL`.** |
| `modDemo.bas` | Runnable smoke-test entry points. |
| `modDvmcp.bas` | Embed live Dataverse data in slide text boxes (see *Embed live data*). |
| `CLAUDE.md` | Deep technical/architecture notes (and AI-assistant guidance). |
| `JsonConverter.bas` | **Not in the repo** — third-party dependency, see below. (Already embedded inside `Presentation1.pptm`.) |

### Dependency: VBA-JSON
VBA has no native JSON parser. `Presentation1.pptm` already contains `JsonConverter`. If you
build a deck from scratch, download `JsonConverter.bas` from
<https://github.com/VBA-tools/VBA-JSON> and import it alongside the other modules.
It needs a reference to **Microsoft Scripting Runtime** (VBE → Tools → References).

---

## Run the demo

First point the deck at a Dataverse environment you can sign in to — see
**[Point it at your own environment](#point-it-at-your-own-environment)**. Then:

1. **Open `Presentation1.pptm`** in PowerPoint and click **Enable Content / Enable Macros**
   when prompted (the modules are already imported).
2. Press **Alt + F11** to open the VBA editor (VBE).
3. In the VBE, confirm **Tools → References → Microsoft Scripting Runtime** is ticked.
4. Put the cursor in `modDemo` → **`Demo_ListTools`** and press **F5**.
   - A browser tab opens at `https://microsoft.com/devicelogin`. A message box shows a
     **user code** — type it in the browser and sign in with a user that has access to the
     Dataverse environment.
   - Back in PowerPoint, click **OK** and wait a few seconds. Expect a message box:
     **"Connected. 15 Dataverse MCP tools:"**.
5. Run **`Demo_CallTool`** (F5) → calls `describe('tables/')` and lists the tables
   (read-only). No second sign-in.
6. **Restart PowerPoint**, reopen the deck, run **`Demo_SilentReconnect`** → reconnects with
   **no browser prompt** (the refresh token was persisted via DPAPI).
   - To force a fresh sign-in: type `SignOut` in the **Immediate** window (Ctrl+G) and
     press Enter, then run `Demo_ListTools` again.

> If you only have the `.bas` files (no deck), follow **[Build a deck from scratch](#build-a-deck-from-scratch)** first, then start at step 4.

---

## Embed live data in a slide (`modDvmcp.bas`)

Author a **command** as the entire text of a text box, run a macro, and the box text is
replaced by live results. The original command is saved on the shape (in its `Tags`), so you
can re-run or edit it later.

1. Insert a text box and type, for example:
   `DVMCP.read_query("select firstname, lastname from contact order by createdon")`
2. Select the box and run **`DVMCP_RefreshSelection`** (Developer → Macros, or F5 in the VBE).
   The text is replaced by the rendered result.
   - Scope variants: **`DVMCP_RefreshSlide`** (whole slide), **`DVMCP_RefreshAll`** (whole deck).
3. To edit the query again, run **`DVMCP_ShowCommands`** — the box reverts to the saved
   command text. Change it, then refresh again. (`…Slide` / `…All` variants exist too.)

**Command syntax:** `DVMCP.tool("argument")` maps the quoted string to the tool's main
parameter (`read_query`→querytext, `search`→query, `describe`→path). Wrap the argument in
double quotes so SQL single-quote literals work, e.g.
`DVMCP.read_query("select name from account where name = 'Contoso'")`.
For any other tool or multiple arguments, pass raw JSON instead:
`DVMCP.search({"query":"contoso","limit":5})`.

If a call fails, the box shows `DVMCP ERROR: …` but keeps the saved command — fix it via
`DVMCP_ShowCommands` and re-run. The macros run silently; status detail is printed to the
**Immediate** window (Ctrl+G in the VBE).

> **Tip — run macros without the VBE.** Add the macros to the Quick Access Toolbar
> (File → Options → Quick Access Toolbar → *Choose commands from: Macros*), or press
> **Alt + F8** in PowerPoint to open the Macros dialog and run them there.

---

## Point it at your own environment

The IDs in this repo are **not secrets** (it's a public-client app), but they are tied to the
author's tenant and Dataverse org. To target your own environment you must (A) do a one-time
Azure / Power Platform setup, (B) edit four constants, and (C) **re-import the changed modules
into the `.pptm`** (editing a `.bas` on disk does *not* change the copy already inside the deck).

### A. One-time setup (Azure + Power Platform)

1. **Entra app registration** (Azure Portal → App registrations → New registration):
   - Authentication → Advanced settings → **Allow public client flows = Yes** (required for
     device code).
   - API permissions → **Dynamics CRM → `mcp.tools`** (delegated) → **Grant admin consent**.
   - No redirect URI is needed for device code flow.
   - Note the **Application (client) ID** and your **Directory (tenant) ID**.
2. **Power Platform Admin Center** (separate from Azure — easy to miss):
   - Environment → Settings → Product → Features → *Dataverse Model Context Protocol* →
     **Advanced Settings** → add a client with your **Application (client) ID**,
     **Is Enabled = Yes**.
   - Without this, `/api/mcp` rejects an otherwise-valid token (401/403).

### B. Edit the four constants

| Constant | File | Line | What to set |
|----------|------|------|-------------|
| `TENANT_ID` | `McpAuth.bas` | 13 | Your Directory (tenant) ID (GUID). |
| `CLIENT_ID` | `McpAuth.bas` | 14 | Your Application (client) ID (GUID). |
| `MCP_SCOPE` | `McpAuth.bas` | 15-16 | `https://<your-org>.crm<NN>.dynamics.com/api/mcp/mcp.tools offline_access openid profile` |
| `MCP_URL`   | `McpClient.bas` | 10 | `https://<your-org>.crm<NN>.dynamics.com/api/mcp` |

Use your environment's real org host (find it in the Power Platform Admin Center, e.g.
`orgXXXXXXXX.crmNN.dynamics.com`). The host in `MCP_SCOPE` and `MCP_URL` must match.

### C. Re-import into the deck

Because `Presentation1.pptm` carries its **own copy** of the modules, you must replace them
after editing:

1. Open `Presentation1.pptm` → **Alt + F11** (VBE).
2. In the Project Explorer, **right-click each module** you changed (`McpAuth`, `McpClient`)
   → **Remove…** → **No** (don't export).
3. **File → Import File…** → select the edited `McpAuth.bas` / `McpClient.bas`.
4. (First time only) Confirm **Tools → References → Microsoft Scripting Runtime** is ticked,
   and that `JsonConverter` is present in the project.
5. **Save** (keep the `.pptm` / macro-enabled format) and run `Demo_ListTools` to verify.

---

## Build a deck from scratch

If you don't want to use the bundled `Presentation1.pptm`:

1. PowerPoint → **Save As → PowerPoint Macro-Enabled Presentation (`.pptm`)**.
2. **Alt + F11** → **File → Import File…** → import all five `.bas` files:
   `McpAuth.bas`, `McpClient.bas`, `modDemo.bas`, `modDvmcp.bas`, and `JsonConverter.bas`
   (download the last one — see *Dependency: VBA-JSON*).
3. **Tools → References → Microsoft Scripting Runtime** → tick it.
4. Edit the constants (see *Point it at your own environment*) if you're not using the
   author's environment.
5. Run `Demo_ListTools` (F5).

---

## Refresh-token persistence (DPAPI)

After the first interactive sign-in, the rotating refresh token is encrypted with the Windows
Data Protection API (**per Windows user + machine**) and stored at
`%APPDATA%\DataverseMcpVba\refresh_<clientid8>.dat`. Subsequent runs — even after restarting
PowerPoint — reconnect silently until the refresh token expires or is revoked. `SignOut`
clears the cache and deletes that file.

- Requires VBA7 (Office 2010+); older Office degrades to one sign-in per session.
- The token does **not** transfer to another computer or Windows user — the first run there
  will require an interactive sign-in. That's expected.
- The `.dat` file is git-ignored and never committed.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---------|--------------------|
| `Compile error` on import | Don't paste module text into the code window — use **File → Import File…** (the `Attribute VB_Name` line is only valid on import). |
| `User-defined type not defined` | **Microsoft Scripting Runtime** reference not ticked, or `JsonConverter` missing. |
| Sign-in fails with **AADSTS70011 / Conditional Access** error | Device code flow is blocked by tenant policy; you'd need auth-code + PKCE (not implemented here). |
| `initialize` returns **401 / 403** with a valid token | The client ID isn't enabled in **Power Platform Admin Center** (setup step A.2), or scope/URL host mismatch. |
| Edited a `.bas` but nothing changed | You edited the file on disk, not the copy inside the `.pptm` — **re-import** (setup step C). |

---

## Tip: prototype outside PowerPoint

`McpAuth` and `McpClient` have **no PowerPoint dependency**. Develop them in **Excel VBA** for
a faster loop, or run the same logic from a `.vbs` script via `cscript.exe` (swap
`MsgBox` / `Debug.Print` for `WScript.Echo`). Only `modDvmcp.bas` is PowerPoint-specific.

---

## Third-party license — VBA-JSON

`JsonConverter.bas` is **VBA-JSON** by Tim Hall, licensed under the **MIT License**
(© 2016 Tim Hall). MIT permits free use, modification, and redistribution — including
**shipping it inside your own `.pptm`** and for commercial use — so bundling it in your
presentation is allowed.

The one obligation: the MIT copyright + permission notice must be **retained in all copies
or substantial portions**. In practice that means **do not strip the header comment block**
at the top of `JsonConverter.bas` (it already declares `(c) Tim Hall … @license MIT` and the
repo URL). As long as that header stays intact in the embedded module, you are compliant.

For belt-and-braces compliance when distributing, also include the full license text — copy
`LICENSE` from <https://github.com/VBA-tools/VBA-JSON> into a `THIRD_PARTY_NOTICES.txt`
alongside your deck. (No notice is needed for the project's own modules, which you author.)
