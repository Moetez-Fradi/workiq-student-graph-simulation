#!/usr/bin/env python3
"""Local Work IQ MCP + Microsoft Graph Education simulator.

This is a deterministic development double, not an authentication provider or
Microsoft service. Run `python3 server.py`, then point a client at
http://127.0.0.1:8787/mcp and http://127.0.0.1:8787/v1.0.
"""
from __future__ import annotations

import json
import os
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlencode, urlparse

HOST = os.environ.get("SIM_HOST", "127.0.0.1")
PORT = int(os.environ.get("SIM_PORT", "8787"))
TOKEN = "sim-access-token-student-001"
AUTH_CODE = "sim-auth-code-student-001"

STUDENT = {
    "id": "student-001",
    "displayName": "Amina Demo",
    "givenName": "Amina",
    "surname": "Demo",
    "mail": "amina.demo@neomind.test",
    "userPrincipalName": "amina.demo@neomind.test",
}
CLASS = {
    "id": "primary-math-2026",
    "displayName": "Primary Mathematics — 2026",
    "description": "Fractions, decimals, and number sense.",
    "classCode": "MATH-PRIMARY-2026",
}
ASSIGNMENT = {
    "id": "fractions-week-1",
    "displayName": "Adding fractions with common denominators",
    "status": "published",
    "dueDateTime": "2026-09-22T16:00:00Z",
}
SCHOOL = {
    "id": "neomind-primary-school",
    "displayName": "NeoMind Primary School",
    "description": "Demo school for the NeoMind primary mathematics curriculum.",
    # Simulator-only extension. The real Graph educationSchool schema has no
    # NeoMind curriculum field; a production integration resolves curriculum
    # through NeoMind's own catalog/enrollment service.
    "neomindCurriculumId": "primary_math",
    "neomindCurriculumLabel": "Primary mathematics",
}
DRIVE_FILES = {
    "value": [
        {
            "id": "fractions-notes",
            "name": "Fractions revision notes.docx",
            "webUrl": "https://neomind.test/drive/fractions-notes",
            "lastModifiedDateTime": "2026-09-10T14:00:00Z",
        }
    ]
}
CALENDAR = {
    "value": [
        {
            "id": "math-review",
            "subject": "Primary Mathematics review",
            "start": {"dateTime": "2026-09-20T15:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-09-20T15:30:00", "timeZone": "UTC"},
        }
    ]
}

PATHS = {
    "/me": ("fetch",),
    "/me/drive/root/search(q='{subject}')": ("fetch",),
    "/me/calendarView?startdatetime={start}&enddatetime={end}": ("fetch",),
    "/education/users/{id}": ("fetch",),
    "/education/schools": ("fetch",),
    "/education/users/{id}/classes": ("fetch",),
    "/education/classes/{id}/members": ("fetch",),
    "/education/classes/{id}/assignments": ("fetch",),
}


def schema_for(path: str) -> dict:
    """Small JSON schemas sufficient for clients to build field mappings."""
    if path == "/me" or path == "/education/users/{id}":
        return {"type": "object", "required": ["id", "displayName"], "properties": {"id": {"type": "string"}, "displayName": {"type": "string"}, "mail": {"type": "string"}}}
    if "assignments" in path:
        item = {"type": "object", "required": ["id", "displayName", "status"], "properties": {"id": {"type": "string"}, "displayName": {"type": "string"}, "status": {"type": "string"}}}
    elif "classes" in path or "members" in path:
        item = {"type": "object", "required": ["id", "displayName"], "properties": {"id": {"type": "string"}, "displayName": {"type": "string"}}}
    elif "calendar" in path:
        item = {"type": "object", "properties": {"id": {"type": "string"}, "subject": {"type": "string"}}}
    else:
        item = {"type": "object", "properties": {"id": {"type": "string"}, "name": {"type": "string"}}}
    return {"type": "object", "properties": {"value": {"type": "array", "items": item}}}


def graph_fetch(path: str) -> tuple[int, dict]:
    """Resolve the supported Graph v1.0-style read paths."""
    base_path = path.split("?", 1)[0]
    if base_path == "/me":
        return 200, STUDENT
    if base_path == "/me/drive/root/search(q='primary_math')" or base_path == "/me/drive/root/search(q='fractions')":
        return 200, DRIVE_FILES
    if base_path == "/me/calendarView":
        return 200, CALENDAR
    if base_path == "/education/users/student-001":
        return 200, {**STUDENT, "primaryRole": "student"}
    if base_path == "/education/schools":
        return 200, {"value": [SCHOOL]}
    if base_path == "/education/users/student-001/classes":
        return 200, {"value": [CLASS]}
    if base_path == "/education/classes/primary-math-2026/members":
        return 200, {"value": [{**STUDENT, "primaryRole": "student"}]}
    if base_path == "/education/classes/primary-math-2026/assignments":
        return 200, {"value": [ASSIGNMENT]}
    return 404, {"error": {"code": "Request_ResourceNotFound", "message": f"Unsupported simulated path: {path}"}}


def mcp_tools() -> list[dict]:
    return [
        {"name": "fetch", "description": "Read Graph entities by relative path.", "inputSchema": {"type": "object", "required": ["entityUrls"], "properties": {"entityUrls": {"type": "array", "items": {"type": "string"}}}}},
        {"name": "search_paths", "description": "Discover supported Graph paths.", "inputSchema": {"type": "object", "required": ["filter"], "properties": {"filter": {"type": "string"}}}},
        {"name": "get_schema", "description": "Retrieve a Graph path JSON schema.", "inputSchema": {"type": "object", "required": ["path", "operationType"], "properties": {"path": {"type": "string"}, "operationType": {"type": "string"}}}},
        {"name": "call_function", "description": "Read a declared parameterised Graph path.", "inputSchema": {"type": "object", "required": ["functionUrl"], "properties": {"functionUrl": {"type": "string"}}}},
    ]


def mcp_call(name: str, arguments: dict) -> dict:
    if name == "search_paths":
        needle = arguments.get("filter", "").lower().replace(".*", "")
        paths = [{"path": path, "operations": list(operations)} for path, operations in PATHS.items() if needle in path.lower()]
        return {"structuredContent": {"paths": paths}, "content": [{"type": "text", "text": json.dumps({"paths": paths})}]}
    if name == "get_schema":
        path = arguments.get("path")
        if not path or arguments.get("operationType") != "fetch":
            return {"isError": True, "content": [{"type": "text", "text": "Only fetch schemas are supported by this read-only simulator."}]}
        schema = schema_for(path)
        return {"structuredContent": schema, "content": [{"type": "text", "text": json.dumps(schema)}]}
    if name == "fetch":
        urls = arguments.get("entityUrls", [])
        if not isinstance(urls, list):
            return {"isError": True, "content": [{"type": "text", "text": "entityUrls must be an array."}]}
        results = [{"statusCode": status, "data": data} for status, data in (graph_fetch(url) for url in urls)]
        return {"structuredContent": {"results": results}, "content": [{"type": "text", "text": json.dumps({"results": results})}]}
    if name == "call_function":
        function_url = arguments.get("functionUrl", "")
        if not function_url.startswith("/me/calendarView?"):
            return {"isError": True, "content": [{"type": "text", "text": "Only /me/calendarView is an allowed simulated function."}]}
        status, data = graph_fetch(function_url)
        return {"structuredContent": {"statusCode": status, "data": data}, "content": [{"type": "text", "text": json.dumps(data)}]}
    return {"isError": True, "content": [{"type": "text", "text": f"Unsupported read-only Work IQ tool: {name}"}]}


class Handler(BaseHTTPRequestHandler):
    server_version = "NeoMindWorkIQSim/0.1"

    def log_message(self, format: str, *args: object) -> None:
        print("[sim] " + format % args)

    def _send(self, status: int, body: dict, *, headers: dict[str, str] | None = None) -> None:
        encoded = json.dumps(body).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        # The Next.js walkthrough runs on a different localhost port.
        self.send_header("Access-Control-Allow-Origin", "http://localhost:3000")
        self.send_header("Access-Control-Allow-Headers", "Authorization, Content-Type")
        self.send_header("Content-Length", str(len(encoded)))
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(encoded)

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "http://localhost:3000")
        self.send_header("Access-Control-Allow-Headers", "Authorization, Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()

    def _json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(length) or b"{}")

    def _authorized(self) -> bool:
        return self.headers.get("Authorization") == f"Bearer {TOKEN}"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self._send(200, {"status": "ok"})
            return
        if parsed.path == "/authorize":
            query = parse_qs(parsed.query)
            redirect_uri = query.get("redirect_uri", [""])[0]
            if not redirect_uri:
                self._send(400, {"error": "invalid_request", "error_description": "redirect_uri is required"})
                return
            separator = "&" if "?" in redirect_uri else "?"
            target = redirect_uri + separator + urlencode({"code": AUTH_CODE, "state": query.get("state", [""])[0]})
            self.send_response(HTTPStatus.FOUND)
            self.send_header("Location", target)
            self.end_headers()
            return
        if parsed.path == "/.well-known/oauth-authorization-server":
            base = f"http://{HOST}:{PORT}"
            self._send(200, {"issuer": base, "authorization_endpoint": base + "/authorize", "token_endpoint": base + "/oauth/token", "response_types_supported": ["code"], "grant_types_supported": ["authorization_code"]})
            return
        if parsed.path.startswith("/v1.0/"):
            if not self._authorized():
                self._send(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Use the simulator bearer token."}})
                return
            status, data = graph_fetch(parsed.path[len("/v1.0"): ] + (("?" + parsed.query) if parsed.query else ""))
            self._send(status, data)
            return
        self._send(404, {"error": "not_found"})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/oauth/token":
            form = parse_qs(self.rfile.read(int(self.headers.get("Content-Length", "0"))).decode())
            if form.get("grant_type", [""])[0] != "authorization_code" or form.get("code", [""])[0] != AUTH_CODE:
                self._send(400, {"error": "invalid_grant"})
                return
            self._send(200, {"access_token": TOKEN, "token_type": "Bearer", "expires_in": 3600, "scope": "User.Read EduRoster.ReadBasic Files.Read"})
            return
        if parsed.path == "/mcp":
            if not self._authorized():
                self._send(401, {"jsonrpc": "2.0", "error": {"code": -32001, "message": "Unauthorized"}, "id": None})
                return
            request = self._json_body()
            request_id = request.get("id")
            if request.get("method") == "initialize":
                result = {"protocolVersion": request.get("params", {}).get("protocolVersion", "2025-03-26"), "capabilities": {"tools": {}}, "serverInfo": {"name": "neomind-workiq-sim", "version": "0.1.0"}}
            elif request.get("method") == "tools/list":
                result = {"tools": mcp_tools()}
            elif request.get("method") == "tools/call":
                params = request.get("params", {})
                result = mcp_call(params.get("name", ""), params.get("arguments", {}))
            else:
                self._send(200, {"jsonrpc": "2.0", "error": {"code": -32601, "message": "Method not found"}, "id": request_id})
                return
            self._send(200, {"jsonrpc": "2.0", "result": result, "id": request_id})
            return
        self._send(404, {"error": "not_found"})


if __name__ == "__main__":
    print(f"NeoMind Work IQ simulation listening at http://{HOST}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
