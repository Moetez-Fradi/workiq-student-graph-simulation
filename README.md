# NeoMind Work IQ + Student Graph simulator

`workiq-student-graph-sim` is a deterministic local double for the read-only subset of Work IQ MCP and Microsoft Graph Education used by the NeoMind proof of concept. It lets an eventual `WorkIQAdapter` exercise OAuth authorization-code-shaped navigation, bearer-token handling, MCP path discovery and schema inspection, and Graph-shaped student, school, class, assignment, file, and calendar data.

It starts with no tenant, Entra app registration, Microsoft 365 account, Copilot Credits, or real student data. The fixture always represents the same demo student and class, so local development and contract tests remain repeatable.

> This is not a Microsoft service, an authentication provider, or a security boundary. Never deploy it, expose it on a network, or use its fixed token outside local development.

## Prerequisites

- Python 3.10 or later (the code uses modern type annotations).
- No third-party packages. See [requirements.txt](requirements.txt) for the intentionally empty dependency policy.

An isolated environment is optional:

```bash
python3 -m venv .venv
. .venv/bin/activate
```

## Start the simulator

From this directory, run:

```bash
python3 server.py
```

The default base URL is `http://127.0.0.1:8787`. Confirm that it is ready:

```bash
curl -s http://127.0.0.1:8787/health
# {"status": "ok"}
```

Use `Ctrl+C` in the terminal running the service to stop it.

### Configuration

The server reads these environment variables at start-up:

| Variable | Default | Meaning |
| --- | --- | --- |
| `SIM_HOST` | `127.0.0.1` | Interface on which the HTTP server listens. Keep the loopback default for local-only use. |
| `SIM_PORT` | `8787` | TCP port for all simulator endpoints. |

For example:

```bash
SIM_PORT=8788 python3 server.py
```

When changing either setting, use the resulting base URL in every client configuration and authorization request.

## Local authentication flow

The server mimics only the outline of delegated authorization. It accepts one fixed authorization code and returns one fixed access token:

| Item | Value |
| --- | --- |
| Authorization code | `sim-auth-code-student-001` |
| Access token | `sim-access-token-student-001` |
| Seeded user ID | `student-001` |

1. Navigate to an authorization URL with a redirect URI owned by your local client:

   ```text
   http://127.0.0.1:8787/authorize?response_type=code&client_id=neomind-demo&redirect_uri=http://127.0.0.1:3000/callback&state=demo
   ```

   The simulator redirects to that URI with `code` and the supplied `state`.

2. Exchange the code for the bearer token:

   ```bash
   curl -s -X POST http://127.0.0.1:8787/oauth/token \
     -H 'Content-Type: application/x-www-form-urlencoded' \
     --data 'grant_type=authorization_code&code=sim-auth-code-student-001'
   ```

3. Include the returned token on Graph and MCP requests:

   ```text
   Authorization: Bearer sim-access-token-student-001
   ```

The discovery document is available at `GET /.well-known/oauth-authorization-server`. Token requests with a different code, and protected requests without the exact bearer token, are rejected.

## Interfaces and examples

| Endpoint | Purpose | Authentication |
| --- | --- | --- |
| `GET /health` | Readiness probe. | No |
| `GET /.well-known/oauth-authorization-server` | Local OAuth metadata. | No |
| `GET /authorize` | Redirects with the fixed authorization code. | No |
| `POST /oauth/token` | Exchanges the fixed code for the fixed token. | No |
| `POST /mcp` | JSON-RPC MCP endpoint. | Bearer token |
| `GET /v1.0/...` | Graph-shaped read endpoints. | Bearer token |

### Microsoft Graph-shaped reads

```bash
curl -s http://127.0.0.1:8787/v1.0/education/users/student-001/classes \
  -H 'Authorization: Bearer sim-access-token-student-001'
```

Supported paths are:

- `/me`
- `/me/drive/root/search(q='primary_math')` and `/me/drive/root/search(q='fractions')`
- `/me/calendarView?startdatetime=...&enddatetime=...`
- `/education/users/student-001`
- `/education/schools`
- `/education/users/student-001/classes`
- `/education/classes/primary-math-2026/members`
- `/education/classes/primary-math-2026/assignments`

Unknown paths return a Graph-style `404` with `Request_ResourceNotFound`. `/education/schools` includes a simulator-only `neomindCurriculumId`; production code must resolve curricula through NeoMind's own catalog or enrollment source instead.

### MCP JSON-RPC endpoint

Initialize the server or list its tools with a JSON-RPC request. For example, read the seeded student and their classes in one `fetch` call:

```bash
curl -s http://127.0.0.1:8787/mcp \
  -H 'Authorization: Bearer sim-access-token-student-001' \
  -H 'Content-Type: application/json' \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"fetch","arguments":{"entityUrls":["/me","/education/users/student-001/classes"]}}}'
```

The read-only tools are:

| Tool | Arguments | Result |
| --- | --- | --- |
| `search_paths` | `filter` | Matching declared relative paths and their operations. |
| `get_schema` | `path`, `operationType: "fetch"` | A small JSON schema for a supported path shape. |
| `fetch` | `entityUrls` array | Ordered `{statusCode, data}` entries, one per requested path. |
| `call_function` | `functionUrl` | The allowed parameterized calendar path only. |

`call_function` accepts only `/me/calendarView?...`. Any missing tool, write operation, or unsupported function responds with an MCP `isError` result.

## Run the tests

The test suite uses only `unittest` and does not need the server to be running:

```bash
python3 -m unittest -v
```

The tests cover seeded Graph responses, Graph-style unknown-resource errors, MCP multi-path fetch ordering, path discovery, and rejection of a mutation tool.

## Design boundaries

The implementation deliberately favors a narrow contract surface over a broad emulation of Microsoft services:

- The HTTP server uses Python's standard-library `http.server`; it is suitable for a local adapter contract test, not production traffic.
- OAuth does not model consent, PKCE, client authentication, refresh tokens, scopes, conditional access, tenant policy, or token validation.
- There are no write routes or MCP mutation tools. That read-only policy is structural: unsupported capabilities do not exist.
- Fixtures contain fabricated demo data only. Do not treat their fields or schemas as a guarantee of Microsoft Graph or Work IQ compatibility.

For real integrations, use Microsoft’s current Work IQ MCP and Microsoft Graph Education documentation as the source of truth; this project is a local contract simulator only.
