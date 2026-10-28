# NeoMind Work IQ + Student Graph simulator

`workiq-student-graph-sim` is a deterministic local double for the three Microsoft services the NeoMind proof of concept talks to: the **Work IQ MCP server**, **Microsoft Graph v1.0** (education, profile, calendar and drive reads), and the **Microsoft Entra ID v2.0** OAuth endpoints that issue tokens for both. It lets an eventual `WorkIQAdapter` exercise the real wire contracts — Entra authorization-code + PKCE, per-resource bearer tokens, MCP Streamable HTTP, Work IQ's tool surface and default tenant policy, and Graph-shaped student, school, class, assignment, file and calendar data — with no tenant, Entra app registration, Microsoft 365 account, Copilot Credits, or real student data.

The fixture always represents the same demo school, class and three people (a student, their teacher and the school's director), so local development and contract tests remain repeatable.

> This is not a Microsoft service, an authentication provider, or a security boundary. Never deploy it, expose it on a network, or use its fixed tokens outside local development.

## Prerequisites

- Python 3.10 or later. No third-party packages — see [requirements.txt](requirements.txt).
- On Windows, [uv](https://docs.astral.sh/uv/) is the easiest way to get a suitable Python (`winget install astral-sh.uv`); `uv run` below fetches one automatically.

## Start the simulator

From this directory:

```bash
python3 server.py                                # Linux / macOS
uv run --no-project --python 3.12 server.py      # any OS, including Windows
```

The default base URL is `http://127.0.0.1:8787`. Confirm that it is ready:

```bash
curl -s http://127.0.0.1:8787/health
# {"status": "ok"}
```

Use `Ctrl+C` to stop it.

### Configuration

| Variable | Default | Meaning |
| --- | --- | --- |
| `SIM_HOST` | `127.0.0.1` | Interface to listen on. Keep the loopback default. |
| `SIM_PORT` | `8787` | TCP port for every endpoint. |
| `SIM_BASE_URL` | `http://{SIM_HOST}:{SIM_PORT}` | URL clients reach the simulator at. It is published in the OIDC metadata, token `iss` claims, the MCP protected-resource metadata and `WWW-Authenticate` challenges. Set it when the bind address isn't what clients use, as in Docker. |
| `SIM_DEFAULT_USER` | the student's UPN | Who signs in when the authorize request has no `login_hint`: `amina.demo@neomind.test`, `demo.teacher@neomind.test` or `principal@neomind.test`. |
| `SIM_WORKIQ_ALLOWED_PREFIXES` | `/me,/users,/sites` | Work IQ tenant-policy path allow-list. The default is Microsoft's documented default; add `/education` to model a tenant whose admin allowed Graph Education paths. |

### Docker

```bash
docker build -t workiq-sim .
docker run --rm -p 127.0.0.1:8787:8787 workiq-sim
```

The image listens on `0.0.0.0:8787` inside the container, so publish it on the loopback interface only, as above. It defaults `SIM_BASE_URL` to `http://127.0.0.1:8787`. If you publish on another host port, set `SIM_BASE_URL` to match (for example `-p 127.0.0.1:9000:8787 -e SIM_BASE_URL=http://127.0.0.1:9000`). The image has a health check against `/health`. Run the tests in it with `docker run --rm workiq-sim python -m unittest -v`.

## Authentication (Entra ID v2.0 shape)

Entra issues an access token for **one resource per request**, so Graph and Work IQ need separate tokens — the simulator enforces this just like the real services (a Graph token on `/mcp`, or a Work IQ token on `/v1.0`, is rejected with 401).

| Item | Value |
| --- | --- |
| Tenant ID | `8f3e2d1c-4b5a-4c6d-9e7f-0a1b2c3d4e5f` (also accepts `common`, `organizations`) |
| Authorization code | `sim-auth-code-student-001` |
| Graph access token (scopes `User.Read EduRoster.ReadBasic EduAssignments.ReadBasic Calendars.Read Files.Read`) | `sim-graph-token-student-001` |
| Work IQ access token (scope `api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask`) | `sim-workiq-token-student-001` |
| Refresh token (issued with `offline_access`) | `sim-refresh-token-student-001` |
| Seeded user (`oid`) | `3f6c1a2b-9d4e-4c8f-a1b2-7e5d9c0f1a23` — `amina.demo@neomind.test` |

The values above are the student's. The teacher's and the director's follow the same pattern with `teacher` or `director` in place of `student` (for example `sim-workiq-token-teacher-001`).

### Who signs in, and their role

Entra signs in whoever picks their account on its page. With no page here, the authorize request's `login_hint` (a UPN, which Entra also accepts) picks the account, then `SIM_DEFAULT_USER`, then the student. An unknown account comes back as `AADSTS50034`.

| Person | UPN | `oid` | `/me` `jobTitle` | `/education/me` `primaryRole` |
| --- | --- | --- | --- | --- |
| Amina Demo, student | `amina.demo@neomind.test` | `3f6c1a2b-9d4e-4c8f-a1b2-7e5d9c0f1a23` | `null` | `student` (with a `student` block) |
| Demo Teacher, class teacher | `demo.teacher@neomind.test` | `7a9c1e3f-5b7d-4f9a-8c1e-3a5c7e9b1d4f` | `Class Teacher` | `teacher` (with a `teacher` block) |
| Demo Principal, director | `principal@neomind.test` | `9e1a3c5f-7d9b-4b1d-a3f5-7c9e1b3d5f8a` | `Director` | `none` |

This is how a real tenant gives a role: Graph's `primaryRole` is only ever `student`, `teacher` or `none`, and a title such as "Director" or "Head Teacher" is the free-text `jobTitle` set by the school's IT or its SIS. There is no single "role" field. Through Work IQ, only `jobTitle` comes back by default: `/me` is allowed and `/education/me` is denied by the default policy below. The director is the school's principal (`principalEmail`) and teaches no class; the teacher is a member of the class and wrote its assignment. Only the student has calendar events and drive files.

1. Send the user to the authorize endpoint (PKCE `code_challenge` optional but verified when present; `response_mode` `query`, `fragment` or `form_post`):

   ```text
   http://127.0.0.1:8787/common/oauth2/v2.0/authorize?client_id=neomind-demo&response_type=code&redirect_uri=http://127.0.0.1:3000/callback&scope=api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask%20offline_access&state=demo
   ```

   The simulator redirects to `redirect_uri` with `code` and `state` (or with `error` / `error_description` carrying an `AADSTS` code, as Entra does).

2. Redeem the code. The `scope` picks the resource — Work IQ scopes (`api://workiq.svc.cloud.microsoft/...`) yield the Work IQ token; Graph scopes (`User.Read`, `https://graph.microsoft.com/.default`, …) yield the Graph token; mixing both is `invalid_scope` (`AADSTS28000`):

   ```bash
   curl -s -X POST http://127.0.0.1:8787/common/oauth2/v2.0/token \
     -d 'grant_type=authorization_code&client_id=neomind-demo&code=sim-auth-code-student-001&redirect_uri=http://127.0.0.1:3000/callback&scope=api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask'
   ```

   Then use the `refresh_token` grant to get the other resource's token, as with Entra. The code can also be redeemed directly without visiting `/authorize` (handy for curl); after an `/authorize` visit it is single-use and bound to that request's `redirect_uri` and PKCE challenge.

3. Send `Authorization: Bearer <token>` on Graph and MCP requests.

Discovery: `GET /{tenant}/v2.0/.well-known/openid-configuration` (Entra's format) and, for MCP clients, `GET /.well-known/oauth-protected-resource/mcp` (RFC 9728), which the `WWW-Authenticate` header on a 401 from `/mcp` points to.

## Interfaces

| Endpoint | Purpose | Auth |
| --- | --- | --- |
| `GET /health` | Readiness probe. | No |
| `GET /{tenant}/v2.0/.well-known/openid-configuration` | Entra OIDC metadata. | No |
| `GET /{tenant}/oauth2/v2.0/authorize` | Authorization-code redirect. | No |
| `POST /{tenant}/oauth2/v2.0/token` | `authorization_code` and `refresh_token` grants; `client_credentials` for NeoMind's app. | No |
| `GET /{tenant}/adminconsent` | The school admin approves NeoMind's app for the tenant. | No |
| `GET /.well-known/oauth-protected-resource/mcp` | MCP protected-resource metadata. | No |
| `POST /mcp` | Work IQ MCP, Streamable HTTP (JSON responses). | Work IQ token |
| `GET /v1.0/...` | Microsoft Graph v1.0 reads. | Graph token |
| `GET /v1.0/sites/...` | The school's curriculum site, read by NeoMind's app. | App token |
| `POST`/`DELETE /sim-admin/...` | Simulator only: what the school's admin approves or withdraws. | No |

### Microsoft Graph reads

```bash
curl -s http://127.0.0.1:8787/v1.0/education/me/classes \
  -H 'Authorization: Bearer sim-graph-token-student-001'
```

Supported paths (case-insensitive, like Graph; `$select` and `$top` work everywhere):

- `/me`, `/users/{id or UPN}`
- `/me/calendarView?startDateTime=...&endDateTime=...` (both required; filtered to the window)
- `/me/drive/root/search(q='...')` (matches any query against the seeded notes), `/me/drive/items/{id}`, `/me/drive/items/{id}/content` (302 to a download URL)
- `/education/me`, `/education/me/classes`, `/education/me/schools`
- `/education/users/{id}`, `/education/users/{id}/classes`, `/education/users/{id}/schools`
- `/education/schools`, `/education/schools/{id}`
- `/education/classes/{id}`, `/education/classes/{id}/members`, `/education/classes/{id}/assignments`, `/education/classes/{id}/assignments/{id}`

Seeded IDs: class `b8d2e4f6-1a3c-4e5b-8d7f-9a0b1c2d3e4f`, school `c4e6a8b0-2d4f-4a6c-8e0b-1d3f5a7c9e2b`, assignment `d5f7b9c1-3e5a-4b7d-9f1c-2e4a6c8e0a3d`.

Errors use Graph's envelope (`error.code`, `error.message`, `error.innerError`): an unknown ID is `404 Request_ResourceNotFound`, an unknown path segment is `400 BadRequest`, a missing calendar window is `400 ErrorInvalidParameter`, and writes are `403 Authorization_RequestDenied` (the student only holds read scopes).

**Mapping to NeoMind:** Graph has no curriculum field. The class's SIS-synced `course.externalId` (`primary_math`) is the realistic hook for choosing NeoMind's catalog, and the student's `student.externalId` (`demo-student-01`) matches NeoMind's fixture `learner_ref`.

### School site (app-only)

NeoMind's curriculum import signs in as itself, not as a user: the `client_credentials` grant with a certificate, against the tenant ID (not `common`). The school's admin approves the app once for the whole tenant and grants it one SharePoint site under `Sites.Selected`.

| Item | Value |
| --- | --- |
| App (client) ID | `6b2d4f8a-1c3e-4a5b-9d7f-2e4c6a8b0d1f`, "NeoMind curriculum import" |
| Scope | `https://graph.microsoft.com/.default` |
| Credential | `client_assertion_type=urn:ietf:params:oauth:client-assertion-type:jwt-bearer` and a `client_assertion` whose header names a certificate (`x5t` or `x5t#S256`), with `iss` and `sub` the client ID, `aud` the token endpoint and `exp` in the future. A client secret is refused (`AADSTS7000215`). |
| App token | `sim-app-token-neomind-import-NNN`, one per sign-in, no refresh token |
| Granted site | `neomind.sharepoint.test,5e7a9c1b-3d5f-4a7b-9c1d-3e5f7a9b1c2d,1a3c5e7f-9b1d-4f3a-8c5e-7a9b1d3f5a6c` (`/sites/curriculum`) |
| Site not granted | `neomind.sharepoint.test,2c4e6a8b-0d2f-4b4d-8f6a-0c2e4a6c8e1f,7b9d1f3a-5c7e-4a9b-8d1f-3b5d7f9a1c4e` (`/sites/staff`) |

The curriculum site's library holds `Maths/Year 4 maths scheme of work.csv` and `Maths/Year 5 maths scheme of work.csv` (columns `year,term,unit,objective`; objectives in the national curriculum's words, plus one the school wrote) and `Maths/Calculation policy.docx`, which is not a real document.

Paths for the app token (a site is its ID or `{hostname}:/sites/{path}`):

- `/sites/{site}`, `/sites/{site}/drive`
- `/sites/{site}/drive/root`, `/sites/{site}/drive/root/children`
- `/sites/{site}/drive/items/{id}`, `/sites/{site}/drive/items/{id}/children` (files carry `@microsoft.graph.downloadUrl`)
- `/sites/{site}/drive/items/{id}/content` (302 to the download URL)

A site that was not granted is `403 accessDenied`, an unknown one `404 itemNotFound`, `/me` is `400 BadRequest` (no user behind the token) and anything else `403 Authorization_RequestDenied`.

Approval, as the school's admin gives it:

- `GET /{tenant}/adminconsent?client_id=...&redirect_uri=...&state=...` approves at once and redirects with `admin_consent=True&tenant=...`.
- `POST` or `DELETE /sim-admin/consent` gives or withdraws the approval. Withdrawn, the app's tokens stop working (`401`) and a new sign-in is `AADSTS7000229`.
- `POST` or `DELETE /sim-admin/sites/{site-id}/grant` grants or revokes one site.
- `POST /sim-admin/reset` goes back to the start: approved, with the curriculum site granted.

```bash
curl -s http://127.0.0.1:8787/v1.0/sites/neomind.sharepoint.test:/sites/curriculum   -H 'Authorization: Bearer sim-app-token-neomind-import-001'
```

### Work IQ MCP endpoint

```bash
curl -s http://127.0.0.1:8787/mcp \
  -H 'Authorization: Bearer sim-workiq-token-student-001' \
  -H 'Content-Type: application/json' -H 'Accept: application/json, text/event-stream' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"fetch","arguments":{"entityUrls":["/me","/education/me/classes"]}}}'
```

The second entry comes back as a `403` policy denial — see the default tenant policy below.

Protocol: JSON-RPC 2.0 over Streamable HTTP; protocol versions `2025-11-25`, `2025-06-18`, `2025-03-26` (negotiated in `initialize`); notifications get `202 Accepted`; `ping` is supported; `GET /mcp` is `405` (no server-initiated stream); foreign `Origin` headers are `403`; unknown tools are JSON-RPC error `-32602`.

The tools and their arguments mirror the [Work IQ MCP tool reference](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/tool-reference):

| Tool | Arguments | Simulator behavior |
| --- | --- | --- |
| `fetch` | `entityUrls` | `structuredContent.results[]` of `{data, statusCode}`, in request order. |
| `fetch_blob` | `path`, `format` | `/me/drive/items/{id}/content` → `{statusCode, contentType, blobName, sizeBytes, base64Content}` in `structuredContent` only. |
| `call_function` | `functionUrl` | `{data, statusCode}` for any supported path (e.g. `/me/calendarView?...`). |
| `search_paths` | `filter` (prefix or regex) | `{paths: [{path, operations}]}` using Graph OpenAPI templates such as `/education/classes/{educationClass-id}/members`. Only `fetch` is listed: there are no write routes. |
| `get_schema` | `path` or `operationIds`, `operationType`, `format` | JSON Schema 2020-12 as text content (`...CollectionResponse` for collections). `operationIds` and `typescript` output aren't simulated. |
| `create_entity`, `update_entity`, `delete_entity`, `do_action` | as documented | Denied by default tenant policy (`isError`). |
| `ask` | `question`, … | `isError`: Microsoft 365 Copilot isn't emulated. |
| `list_agents` | none | The built-in Microsoft Copilot agent. |

**Default tenant policy**, as documented by Microsoft: entity tools may only touch paths under `/me/`, `/users/` and `/sites/`; `/authentication/` and `/servicePrincipals/` are blocked; `$skip`/`$skiptoken` are blocked; collection reads get `$top=25` unless specified, capped at 100; mutations are denied. **`/education/...` paths are therefore denied through Work IQ by default** — read them from Graph directly, or set `SIM_WORKIQ_ALLOWED_PREFIXES` to model a tenant that allows them. A denied `fetch` path comes back as its own `{statusCode: 403, data: {error: {code: "PolicyDenied", ...}}}` entry.

## Run the tests

The suite uses only `unittest`; it starts its own server on a random port:

```bash
python3 -m unittest -v                                   # Linux / macOS
uv run --no-project --python 3.12 -m unittest -v         # any OS
```

## Conformance

Sources: the Work IQ MCP [tool reference](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/tool-reference), [policy governance](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/policy-governance-mcp) and [permissions](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/permissions) pages; the Microsoft Graph v1.0 reference for each resource; the Microsoft identity platform v2.0 authorization-code documentation; and the MCP specification (tools, Streamable HTTP, authorization).

Where Microsoft doesn't document a shape, the simulator makes an explicit, commented assumption:

- The wire shape of a Work IQ policy denial (a Graph-style 403 `PolicyDenied` body per `fetch` entry, `isError` text for denied tools).
- `id_token`s are unsigned (`alg: none`) and the JWKS is empty; real Entra tokens are RS256 JWTs, and real access tokens are JWTs rather than opaque strings.
- A client assertion's signature isn't checked (there are no keys); its header must name a certificate and its claims must be right. The wording for a wrong `aud` is assumed.
- Withdrawing admin consent ends the app's tokens at once; in Entra they last until they expire, about an hour. `/sim-admin/` stands in for the Entra admin centre and for `POST /sites/{id}/permissions`.
- Like Entra, the OIDC metadata doesn't advertise `code_challenge_methods_supported`, even though `S256` works. Some strict MCP clients refuse to proceed without it — the same friction they'd hit against Entra.

## Design boundaries

- Standard-library `http.server`: suitable for local contract tests, not production traffic.
- OAuth doesn't model user consent, conditional access, token expiry, or signature validation. Admin consent for NeoMind's app is modelled, without a page.
- There are no write routes. Mutation tools exist (the real server lists them) but are always denied, as by Work IQ's default policy.
- Fixtures contain fabricated demo data only.

For real integrations, Microsoft's current Work IQ MCP and Microsoft Graph documentation remain the source of truth.
