#!/usr/bin/env python3
"""Local Work IQ MCP + Microsoft Graph Education + Entra ID simulator.

This is a deterministic development double, not an authentication provider or
Microsoft service. Run `python3 server.py`, then point a client at
http://127.0.0.1:8787/mcp (Work IQ MCP), http://127.0.0.1:8787/v1.0 (Graph)
and http://127.0.0.1:8787/{tenant}/oauth2/v2.0 (Entra ID).

Wire shapes follow Microsoft's published contracts (see README "Conformance").
Where Microsoft doesn't document a shape, the code says "simulator assumption".
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import time
import uuid
from datetime import datetime, timezone
from html import escape
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlencode, urlparse

HOST = os.environ.get("SIM_HOST", "127.0.0.1")
PORT = int(os.environ.get("SIM_PORT", "8787"))
BASE_URL = f"http://{HOST}:{PORT}"
GRAPH_CONTEXT = "https://graph.microsoft.com/v1.0/$metadata#"

# --- Identity fixtures -------------------------------------------------------

TENANT_ID = "8f3e2d1c-4b5a-4c6d-9e7f-0a1b2c3d4e5f"
TENANT_ALIASES = {"common", "organizations", TENANT_ID, "neomind.test"}
AUTH_CODE = "sim-auth-code-student-001"
REFRESH_TOKEN = "sim-refresh-token-student-001"
# Entra issues one access token per resource: Graph and Work IQ reject each
# other's tokens (wrong audience), exactly like the real services.
GRAPH_TOKEN = "sim-graph-token-student-001"
WORKIQ_TOKEN = "sim-workiq-token-student-001"
WORKIQ_RESOURCE = "api://workiq.svc.cloud.microsoft"
WORKIQ_SCOPE = f"{WORKIQ_RESOURCE}/WorkIQAgent.Ask"
GRAPH_SCOPES = ("User.Read", "EduRoster.ReadBasic", "EduAssignments.ReadBasic", "Calendars.Read", "Files.Read")
OIDC_SCOPES = ("openid", "profile", "email", "offline_access")

STUDENT_ID = "3f6c1a2b-9d4e-4c8f-a1b2-7e5d9c0f1a23"
CLASS_ID = "b8d2e4f6-1a3c-4e5b-8d7f-9a0b1c2d3e4f"
SCHOOL_ID = "c4e6a8b0-2d4f-4a6c-8e0b-1d3f5a7c9e2b"
ASSIGNMENT_ID = "d5f7b9c1-3e5a-4b7d-9f1c-2e4a6c8e0a3d"
UPN = "amina.demo@neomind.test"

# --- Graph fixtures (v1.0 resource shapes) -----------------------------------

STUDENT = {  # microsoft.graph.user
    "businessPhones": [],
    "displayName": "Amina Demo",
    "givenName": "Amina",
    "jobTitle": None,
    "mail": UPN,
    "mobilePhone": None,
    "officeLocation": None,
    "preferredLanguage": "en-GB",
    "surname": "Demo",
    "userPrincipalName": UPN,
    "id": STUDENT_ID,
}
EDUCATION_STUDENT = {  # microsoft.graph.educationUser
    "id": STUDENT_ID,
    "accountEnabled": True,
    "displayName": "Amina Demo",
    "givenName": "Amina",
    "surname": "Demo",
    "mail": UPN,
    "userPrincipalName": UPN,
    "userType": "Member",
    "primaryRole": "student",
    "onPremisesInfo": {"immutableId": None},
    # externalId is the SIS key; it matches NeoMind's fixture learner_ref.
    "student": {
        "birthDate": None,
        "externalId": "demo-student-01",
        "gender": None,
        "grade": "4",
        "graduationYear": "2032",
        "studentNumber": "0001",
    },
}
CLASS = {  # microsoft.graph.educationClass
    "id": CLASS_ID,
    "displayName": "Primary Mathematics — 2026",
    "description": "Fractions, decimals, and number sense.",
    "mailNickname": "PrimaryMath2026",
    "classCode": "MATH-PRIMARY-2026",
    "externalId": "primary-math-2026",
    "externalName": "Primary Mathematics",
    "externalSource": "sis",
    "externalSourceDetail": "NeoMind demo SIS",
    "grade": "4",
    # Real Graph has no curriculum field; a SIS-synced course externalId is the
    # realistic hook for mapping a class onto NeoMind's catalog.
    "course": {
        "code": "MATH4",
        "courseNumber": "MATH-4",
        "description": "Primary mathematics curriculum.",
        "displayName": "Primary Mathematics",
        "externalId": "primary_math",
        "subject": "Mathematics",
    },
    "term": {
        "displayName": "2026-27",
        "externalId": "2026-27",
        "startDate": "2026-09-01",
        "endDate": "2027-07-16",
    },
    "createdBy": {"user": {"displayName": "NeoMind SIS sync", "id": None}},
}
SCHOOL = {  # microsoft.graph.educationSchool
    "id": SCHOOL_ID,
    "displayName": "NeoMind Primary School",
    "description": "Demo school for the NeoMind primary mathematics curriculum.",
    "externalSource": "sis",
    "externalSourceDetail": "NeoMind demo SIS",
    "principalEmail": "principal@neomind.test",
    "principalName": "Demo Principal",
    "externalPrincipalId": "principal-01",
    "lowestGrade": "1",
    "highestGrade": "6",
    "schoolNumber": "0421",
    "externalId": "neomind-primary",
    "phone": None,
    "fax": None,
    "address": {"city": None, "countryOrRegion": None, "postalCode": None, "state": None, "street": None},
    "createdBy": {"user": {"displayName": "NeoMind SIS sync", "id": None}},
}
ASSIGNMENT = {  # microsoft.graph.educationAssignment, as a student sees it
    "id": ASSIGNMENT_ID,
    "classId": CLASS_ID,
    "displayName": "Adding fractions with common denominators",
    "status": "assigned",
    "instructions": {"content": "Complete the fraction addition practice set.", "contentType": "text"},
    "assignDateTime": None,
    "assignedDateTime": "2026-09-15T08:00:00Z",
    "dueDateTime": "2026-09-22T16:00:00Z",
    "closeDateTime": None,
    "allowLateSubmissions": True,
    "allowStudentsToAddResourcesToSubmission": True,
    "addedStudentAction": "none",
    "addToCalendarAction": "none",
    "assignTo": {"@odata.type": "#microsoft.graph.educationAssignmentClassRecipient"},
    "grading": {"@odata.type": "#microsoft.graph.educationAssignmentPointsGradeType", "maxPoints": 10},
    "languageTag": "en-GB",
    "createdDateTime": "2026-09-14T17:00:00Z",
    "lastModifiedDateTime": "2026-09-15T08:00:00Z",
    "createdBy": {"user": {"displayName": "Demo Teacher", "id": None}},
    "lastModifiedBy": {"user": {"displayName": "Demo Teacher", "id": None}},
    "webUrl": "https://teams.microsoft.com/l/entity/66aeee93-507d-479a-a3ef-8f494af43945/classroom",
    "resourcesFolderUrl": None,
    "feedbackResourcesFolderUrl": None,
    "notificationChannelUrl": None,
    "moduleUrl": None,
}
NOTES_CONTENT = b"Fractions revision notes\n\nSame denominator: add the numerators, keep the denominator.\n"
DRIVE_ITEM = {  # microsoft.graph.driveItem
    "id": "01NEOMINDSIMFRACTIONSNOTES00001",
    "name": "Fractions revision notes.txt",
    "size": len(NOTES_CONTENT),
    "webUrl": "https://neomind-my.sharepoint.test/personal/amina_demo/Documents/Fractions%20revision%20notes.txt",
    "createdDateTime": "2026-09-10T13:00:00Z",
    "lastModifiedDateTime": "2026-09-10T14:00:00Z",
    "file": {"mimeType": "text/plain"},
    "parentReference": {"driveId": "b!neomind-sim-drive", "driveType": "business", "id": "01NEOMINDSIMDRIVEROOT000000000"},
}
DRIVE_ITEMS = {DRIVE_ITEM["id"].lower(): DRIVE_ITEM}
# Drive search is full-text in Graph; this is the text the simulator matches.
DRIVE_SEARCH_TEXT = {DRIVE_ITEM["id"]: "fractions revision notes primary_math adding fractions common denominators"}
EVENTS = [  # microsoft.graph.event
    {
        "id": "AAMkADNlb21pbmQtc2ltAEYAAAAAbWF0aC1yZXZpZXc=",
        "subject": "Primary Mathematics review",
        "isAllDay": False,
        "start": {"dateTime": "2026-09-20T15:00:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2026-09-20T15:30:00.0000000", "timeZone": "UTC"},
        "location": {"displayName": "Room 4B", "locationType": "default", "uniqueId": "Room 4B", "uniqueIdType": "private"},
        "organizer": {"emailAddress": {"name": "Primary Mathematics — 2026", "address": "PrimaryMath2026@neomind.test"}},
    }
]

# --- Graph router ------------------------------------------------------------


def _now() -> datetime:
    return datetime.now(timezone.utc)


def graph_error(status: int, code: str, message: str) -> tuple[int, dict]:
    request_id = str(uuid.uuid4())
    inner = {"date": _now().strftime("%Y-%m-%dT%H:%M:%S"), "request-id": request_id, "client-request-id": request_id}
    return status, {"error": {"code": code, "message": message, "innerError": inner}}


def _not_found(resource_id: str) -> tuple[int, dict]:
    return graph_error(
        404,
        "Request_ResourceNotFound",
        f"Resource '{resource_id}' does not exist or one of its queried reference-property objects are not present.",
    )


def _entity(context: str, entity: dict) -> tuple[int, dict]:
    return 200, {"@odata.context": GRAPH_CONTEXT + context, **entity}


def _collection(context: str, items: list[dict]) -> tuple[int, dict]:
    return 200, {"@odata.context": GRAPH_CONTEXT + context, "value": items}


def _is_student(user_id: str | None) -> bool:
    """`None` is the /me alias; users can also be addressed by UPN."""
    return user_id is None or user_id.lower() in (STUDENT_ID, UPN)


def _parse_datetime(value: str) -> datetime:
    """ISO 8601 with Graph's 7-digit fractions; naive values are UTC."""
    value = re.sub(r"(\.\d{6})\d+", r"\1", value.strip().replace("Z", "+00:00"))
    parsed = datetime.fromisoformat(value)
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def _user(match: re.Match, query: dict) -> tuple[int, dict]:
    return _entity("users/$entity", STUDENT) if _is_student(match["user"]) else _not_found(match["user"])


def _education_user(match: re.Match, query: dict) -> tuple[int, dict]:
    if not _is_student(match["user"]):
        return _not_found(match["user"])
    return _entity("education/users/$entity", EDUCATION_STUDENT)


def _user_classes(match: re.Match, query: dict) -> tuple[int, dict]:
    if not _is_student(match["user"]):
        return _not_found(match["user"])
    return _collection("education/classes", [CLASS])


def _user_schools(match: re.Match, query: dict) -> tuple[int, dict]:
    if not _is_student(match["user"]):
        return _not_found(match["user"])
    return _collection("education/schools", [SCHOOL])


def _schools(match: re.Match, query: dict) -> tuple[int, dict]:
    return _collection("education/schools", [SCHOOL])


def _school(match: re.Match, query: dict) -> tuple[int, dict]:
    if match["school"].lower() != SCHOOL_ID:
        return _not_found(match["school"])
    return _entity("education/schools/$entity", SCHOOL)


def _class(match: re.Match, query: dict) -> tuple[int, dict]:
    if match["cls"].lower() != CLASS_ID:
        return _not_found(match["cls"])
    return _entity("education/classes/$entity", CLASS)


def _class_members(match: re.Match, query: dict) -> tuple[int, dict]:
    if match["cls"].lower() != CLASS_ID:
        return _not_found(match["cls"])
    return _collection("Collection(microsoft.graph.educationUser)", [EDUCATION_STUDENT])


def _class_assignments(match: re.Match, query: dict) -> tuple[int, dict]:
    if match["cls"].lower() != CLASS_ID:
        return _not_found(match["cls"])
    return _collection(f"education/classes('{CLASS_ID}')/assignments", [ASSIGNMENT])


def _class_assignment(match: re.Match, query: dict) -> tuple[int, dict]:
    if match["cls"].lower() != CLASS_ID:
        return _not_found(match["cls"])
    if match["assignment"].lower() != ASSIGNMENT_ID:
        return _not_found(match["assignment"])
    return _entity(f"education/classes('{CLASS_ID}')/assignments/$entity", ASSIGNMENT)


def _calendar_view(match: re.Match, query: dict) -> tuple[int, dict]:
    params = {key.lower(): values[-1] for key, values in query.items()}
    if not params.get("startdatetime") or not params.get("enddatetime"):
        return graph_error(
            400,
            "ErrorInvalidParameter",
            "This request requires a time window specified by the query string parameters StartDateTime and EndDateTime.",
        )
    try:
        start, end = _parse_datetime(params["startdatetime"]), _parse_datetime(params["enddatetime"])
    except ValueError:
        return graph_error(400, "ErrorInvalidParameter", "StartDateTime and EndDateTime must be ISO 8601 date/time values.")
    events = [
        event
        for event in EVENTS
        if _parse_datetime(event["start"]["dateTime"]) < end and _parse_datetime(event["end"]["dateTime"]) > start
    ]
    return _collection(f"users('{STUDENT_ID}')/calendarView", events)


def _drive_search(match: re.Match, query: dict) -> tuple[int, dict]:
    needle = match["q"].lower()
    items = [item for item in DRIVE_ITEMS.values() if needle in DRIVE_SEARCH_TEXT[item["id"]]]
    return _collection("Collection(driveItem)", items)


def _drive_item(match: re.Match, query: dict) -> tuple[int, dict]:
    item = DRIVE_ITEMS.get(match["item"].lower())
    if item is None:
        return graph_error(404, "itemNotFound", "The resource could not be found.")
    return _entity(f"users('{STUDENT_ID}')/drive/items/$entity", item)


def _route(pattern: str) -> re.Pattern:
    return re.compile(pattern, re.IGNORECASE)


USER = r"(?:me|users/(?P<user>[^/]+))"
GRAPH_ROUTES = [
    (_route(r"/me|/users/(?P<user>[^/]+)"), _user),
    (_route(r"/me/calendarView"), _calendar_view),
    (_route(r"/me/drive/root/search\(q='(?P<q>[^']*)'\)"), _drive_search),
    (_route(r"/me/drive/items/(?P<item>[^/]+)"), _drive_item),
    (_route(rf"/education/{USER}"), _education_user),
    (_route(rf"/education/{USER}/classes"), _user_classes),
    (_route(rf"/education/{USER}/schools"), _user_schools),
    (_route(r"/education/schools"), _schools),
    (_route(r"/education/schools/(?P<school>[^/]+)"), _school),
    (_route(r"/education/classes/(?P<cls>[^/]+)"), _class),
    (_route(r"/education/classes/(?P<cls>[^/]+)/members"), _class_members),
    (_route(r"/education/classes/(?P<cls>[^/]+)/assignments"), _class_assignments),
    (_route(r"/education/classes/(?P<cls>[^/]+)/assignments/(?P<assignment>[^/]+)"), _class_assignment),
]


def _apply_query_options(body: dict, query: dict) -> tuple[int, dict]:
    """Supports the $top and $select OData options on any route."""
    options = {key.lower(): values[-1] for key, values in query.items()}
    fields = [field.strip().lower() for field in options.get("$select", "").split(",") if field.strip()]

    def select(entity: dict) -> dict:
        if not fields:
            return entity
        return {key: value for key, value in entity.items() if key.startswith("@odata") or key.lower() in fields}

    if "value" not in body:
        return 200, select(body)
    items = body["value"]
    if "$top" in options:
        if not options["$top"].isdigit():
            return graph_error(400, "BadRequest", f"Invalid value '{options['$top']}' for $top query option.")
        items = items[: int(options["$top"])]
    return 200, {**body, "value": [select(item) for item in items]}


def graph_fetch(path: str) -> tuple[int, dict]:
    """Resolve a Graph v1.0 relative read path (without the /v1.0 prefix)."""
    parsed = urlparse(path)
    route_path = unquote(parsed.path).rstrip("/") or "/"
    query = parse_qs(parsed.query, keep_blank_values=True)
    for pattern, handler in GRAPH_ROUTES:
        match = pattern.fullmatch(route_path)
        if match:
            status, body = handler(match, query)
            return _apply_query_options(body, query) if status == 200 else (status, body)
    segment = route_path.rsplit("/", 1)[-1]
    return graph_error(400, "BadRequest", f"Resource not found for the segment '{segment}'.")


def drive_content(path: str) -> dict | None:
    """The bytes behind /me/drive/items/{id}/content, or None if unknown."""
    match = re.fullmatch(r"/me/drive/items/([^/]+)/content", unquote(urlparse(path).path), re.IGNORECASE)
    item = DRIVE_ITEMS.get(match[1].lower()) if match else None
    if item is None:
        return None
    return {"name": item["name"], "contentType": item["file"]["mimeType"], "content": NOTES_CONTENT}


# --- Work IQ policy ----------------------------------------------------------

# Microsoft's documented defaults. Tenants can widen the allowed prefixes, so
# SIM_WORKIQ_ALLOWED_PREFIXES="/me,/users,/sites,/education" models a tenant
# whose admin allowed Graph Education paths.
WORKIQ_ALLOWED_PREFIXES = tuple(
    prefix.strip().rstrip("/")
    for prefix in os.environ.get("SIM_WORKIQ_ALLOWED_PREFIXES", "/me,/users,/sites").split(",")
    if prefix.strip()
)
WORKIQ_BLOCKED_SEGMENTS = ("/authentication/", "/servicePrincipals/")
WORKIQ_DEFAULT_TOP = 25
WORKIQ_MAX_TOP = 100
MUTATION_TOOLS = ("create_entity", "update_entity", "delete_entity", "do_action")


def policy_denial(url: str) -> str | None:
    """Why the default tenant policy blocks `url`, or None if it is allowed."""
    parsed = urlparse(url)
    path = unquote(parsed.path)
    if not path.startswith("/"):
        return "resource paths must be relative and start with '/'"
    bounded = path.lower() + "/"
    for segment in WORKIQ_BLOCKED_SEGMENTS:
        if segment.lower() in bounded:
            return f"the path segment '{segment}' is blocked"
    if not any(bounded.startswith(prefix.lower() + "/") for prefix in WORKIQ_ALLOWED_PREFIXES):
        return f"'{path}' is outside the allowed path prefixes ({', '.join(p + '/' for p in WORKIQ_ALLOWED_PREFIXES)})"
    options = {key.lower() for key in parse_qs(parsed.query, keep_blank_values=True)}
    if options & {"$skip", "$skiptoken"}:
        return "the $skip and $skiptoken query parameters are blocked"
    return None


def policy_error(reason: str) -> dict:
    # Simulator assumption: Microsoft documents that policy denials happen but
    # not their wire shape, so this mirrors a Graph-style 403 error body.
    return {"error": {"code": "PolicyDenied", "message": f"Request denied by Work IQ tenant policy: {reason}."}}


def workiq_graph(url: str) -> tuple[int, dict]:
    """A Graph read as Work IQ performs it: policy check, then a bounded $top."""
    reason = policy_denial(url)
    if reason:
        return 403, policy_error(reason)
    parsed = urlparse(url)
    query = parse_qs(parsed.query, keep_blank_values=True)
    top = next((values[-1] for key, values in query.items() if key.lower() == "$top"), None)
    query = {key: values for key, values in query.items() if key.lower() != "$top"}
    top_value = min(int(top), WORKIQ_MAX_TOP) if top and top.isdigit() else WORKIQ_DEFAULT_TOP
    query["$top"] = [str(top_value)]
    return graph_fetch(parsed.path + "?" + urlencode(query, doseq=True))


# --- Work IQ MCP tools -------------------------------------------------------

SUPPORTED_PROTOCOL_VERSIONS = ("2025-11-25", "2025-06-18", "2025-03-26")
BUILTIN_AGENT = {"agentId": "bizchat-as-gpt-scenario", "name": "Microsoft Copilot", "provider": "Microsoft"}

# Graph v1.0 path templates this simulator serves (using Graph's OpenAPI
# parameter names), mapped to their response type and whether it is a
# collection. Only `fetch` is listed: the simulator has no write routes.
PATHS = {
    "/me": ("microsoft.graph.user", False),
    "/users/{user-id}": ("microsoft.graph.user", False),
    "/me/calendarView": ("microsoft.graph.event", True),
    "/me/drive/root/search(q='{q}')": ("microsoft.graph.driveItem", True),
    "/me/drive/items/{driveItem-id}": ("microsoft.graph.driveItem", False),
    "/education/me": ("microsoft.graph.educationUser", False),
    "/education/me/classes": ("microsoft.graph.educationClass", True),
    "/education/me/schools": ("microsoft.graph.educationSchool", True),
    "/education/users/{educationUser-id}": ("microsoft.graph.educationUser", False),
    "/education/users/{educationUser-id}/classes": ("microsoft.graph.educationClass", True),
    "/education/users/{educationUser-id}/schools": ("microsoft.graph.educationSchool", True),
    "/education/schools": ("microsoft.graph.educationSchool", True),
    "/education/schools/{educationSchool-id}": ("microsoft.graph.educationSchool", False),
    "/education/classes/{educationClass-id}": ("microsoft.graph.educationClass", False),
    "/education/classes/{educationClass-id}/members": ("microsoft.graph.educationUser", True),
    "/education/classes/{educationClass-id}/assignments": ("microsoft.graph.educationAssignment", True),
    "/education/classes/{educationClass-id}/assignments/{educationAssignment-id}": ("microsoft.graph.educationAssignment", False),
}
TYPE_SAMPLES = {
    "microsoft.graph.user": STUDENT,
    "microsoft.graph.event": EVENTS[0],
    "microsoft.graph.driveItem": DRIVE_ITEM,
    "microsoft.graph.educationUser": EDUCATION_STUDENT,
    "microsoft.graph.educationClass": CLASS,
    "microsoft.graph.educationSchool": SCHOOL,
    "microsoft.graph.educationAssignment": ASSIGNMENT,
}


class JsonRpcError(Exception):
    def __init__(self, code: int, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


def _json_schema_of(value: object) -> dict:
    if isinstance(value, bool):
        return {"type": "boolean"}
    if isinstance(value, int):
        return {"type": "integer"}
    if isinstance(value, dict):
        return {"type": "object", "properties": {key: _json_schema_of(item) for key, item in value.items()}}
    if isinstance(value, list):
        return {"type": "array", "items": _json_schema_of(value[0]) if value else {}}
    if value is None:
        return {"type": ["string", "null"]}
    return {"type": "string"}


def schema_for(template: str) -> dict:
    """A JSON Schema (2020-12) for a path's fetch response, in get_schema's shape."""
    type_name, is_collection = PATHS[template]
    schema: dict = {"$schema": "https://json-schema.org/draft/2020-12/schema"}
    if is_collection:
        schema["title"] = f"{type_name}CollectionResponse"
        schema["type"] = "object"
        schema["properties"] = {"value": {"type": "array", "items": {"$ref": f"#/$defs/{type_name}"}}}
    else:
        schema["title"] = type_name
        schema["$ref"] = f"#/$defs/{type_name}"
    schema["$defs"] = {type_name: _json_schema_of(TYPE_SAMPLES[type_name])}
    return schema


def _template_for(path: str) -> str | None:
    """Map a template or concrete path onto one of PATHS."""
    if path in PATHS:
        return path
    concrete = unquote(urlparse(path).path).rstrip("/")
    for template in PATHS:
        pattern = re.escape(template).replace(r"\{", "{").replace(r"\}", "}")
        pattern = re.sub(r"\{[^}]+\}", r"[^/]+", pattern)
        if re.fullmatch(pattern, concrete, re.IGNORECASE):
            return template
    return None


def _string(description: str) -> dict:
    return {"type": "string", "description": description}


_AGENT_ID = _string("Reserved for future use.")
_JSON_BODY = _string("The entity data, matching the resource schema. This property must be a JSON-encoded string, not a JSON object.")
MCP_TOOLS = [
    {
        "name": "fetch",
        "description": "Reads one or more entities by resource path. Supports fetching multiple paths in parallel.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "entityUrls": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "The relative resource paths to fetch (for example, /me/messages, /me/events/{event-id})",
                },
                "agentId": _AGENT_ID,
            },
            "required": ["entityUrls"],
        },
    },
    {
        "name": "fetch_blob",
        "description": "Fetches binary content, such as documents, images, and Office files from a relative Work IQ path. The content is returned as a Base64-encoded string with metadata.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": _string("The relative path to the binary content (for example, /drives/{drive-id}/items/{item-id}/content)"),
                "format": _string("File conversion format (for example, pdf)."),
                "agentId": _AGENT_ID,
            },
            "required": ["path"],
        },
    },
    {
        "name": "create_entity",
        "description": "Creates a new entity in a collection.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "parentUrl": _string("The relative resource path for the collection (for example, /me/events, /me/messages)."),
                "jsonBody": _JSON_BODY,
                "agentId": _AGENT_ID,
            },
            "required": ["parentUrl", "jsonBody"],
        },
    },
    {
        "name": "update_entity",
        "description": "Updates an existing entity.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "entityUrl": _string("The relative path to the entity (for example, /me/messages/{message-id})."),
                "jsonBody": _JSON_BODY,
                "agentId": _AGENT_ID,
            },
            "required": ["entityUrl", "jsonBody"],
        },
    },
    {
        "name": "delete_entity",
        "description": "Deletes an entity.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "entityUrl": _string("The relative resource path to the entity to delete (for example, /me/messages/{id})"),
                "agentId": _AGENT_ID,
            },
            "required": ["entityUrl"],
        },
    },
    {
        "name": "do_action",
        "description": "Executes a side-effect action such as sending mail, copying, or moving items.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "actionUrl": _string("The relative path for the action (for example, /me/messages/{message-id}/send)."),
                "jsonBody": _JSON_BODY,
                "agentId": _AGENT_ID,
            },
            "required": ["actionUrl"],
        },
    },
    {
        "name": "call_function",
        "description": "Calls a Microsoft Graph function to compute derived data such as schedules, deltas, or search results.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "functionUrl": _string("The relative path for the function (for example, /me/calendarview)."),
                "agentId": _AGENT_ID,
            },
            "required": ["functionUrl"],
        },
    },
    {
        "name": "ask",
        "description": "Asks Microsoft 365 Copilot (or a specific agent) a natural-language question about the user's data.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "question": _string("The natural-language question to ask"),
                "agentId": _string("The ID of a specific agent to route the question to. If omitted, defaults to the built-in Microsoft 365 Copilot agent."),
                "fileUrls": {"type": "array", "items": {"type": "string"}, "description": "An array of OneDrive or SharePoint file URLs to use as context."},
                "conversationId": _string("A conversation ID from an existing conversation."),
                "timeZone": _string("An IANA time zone identifier matching the user's current UTC offset."),
            },
            "required": ["question"],
        },
    },
    {
        "name": "list_agents",
        "description": "Lists available agents that you can use with the ask tool.",
        "inputSchema": {"type": "object", "additionalProperties": False},
    },
    {
        "name": "get_schema",
        "description": "Retrieves the OpenAPI schema for a specific operation, identified by path and operation type.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "operationIds": _string("The operation ID of the API to get the schema for (for example, me.CreateMessages). Use operationIds or path, not both."),
                "path": _string("The API path to get the schema for (for example, /me/messages). Use operationIds or path, not both."),
                "operationType": {"type": "string", "enum": ["fetch", "create", "update"], "description": "The operation type."},
                "format": {"type": "string", "enum": ["jsonschema", "typescript"], "description": "The desired output format. If omitted, JSON schema is returned."},
                "backend": _AGENT_ID,
                "agentId": _AGENT_ID,
            },
            "required": ["operationType"],
        },
    },
    {
        "name": "search_paths",
        "description": "Searches available API paths by prefix or regex filter.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "filter": _string("A prefix or regex pattern to search for matching API paths (for example, messages or .*calendar.*)"),
                "backend": _AGENT_ID,
                "agentId": _AGENT_ID,
            },
            "required": ["filter"],
        },
    },
]
MCP_TOOL_NAMES = {tool["name"] for tool in MCP_TOOLS}


def _text_result(text: str, *, structured: object = None, is_error: bool = False) -> dict:
    result: dict = {"content": [{"type": "text", "text": text}]}
    if structured is not None:
        result["structuredContent"] = structured
    if is_error:
        result["isError"] = True
    return result


def _json_result(structured: dict) -> dict:
    return _text_result(json.dumps(structured), structured=structured)


def _tool_error(message: str) -> dict:
    return _text_result(message, is_error=True)


def _string_argument(arguments: dict, name: str) -> str | None:
    value = arguments.get(name)
    return value if isinstance(value, str) and value else None


def mcp_call(name: str, arguments: dict) -> dict:
    """Run one tools/call. Unknown tools are a protocol error, per the MCP spec."""
    if name not in MCP_TOOL_NAMES:
        raise JsonRpcError(-32602, f"Unknown tool: {name}")

    if name in MUTATION_TOOLS:
        return _tool_error(
            f"Request denied by Work IQ tenant policy: {name} is a mutation operation. Mutation operations "
            "aren't allowed by default; a tenant administrator can enable them in the Microsoft 365 admin center."
        )

    if name == "fetch":
        urls = arguments.get("entityUrls")
        if not isinstance(urls, list) or not urls or not all(isinstance(url, str) for url in urls):
            return _tool_error("entityUrls must be a non-empty array of relative resource paths.")
        results = [{"data": data, "statusCode": status} for status, data in map(workiq_graph, urls)]
        return _json_result({"results": results})

    if name == "call_function":
        url = _string_argument(arguments, "functionUrl")
        if url is None:
            return _tool_error("functionUrl is required.")
        status, data = workiq_graph(url)
        return _json_result({"data": data, "statusCode": status})

    if name == "fetch_blob":
        path = _string_argument(arguments, "path")
        if path is None:
            return _tool_error("path is required.")
        reason = policy_denial(path)
        if reason:
            return _tool_error(f"Request denied by Work IQ tenant policy: {reason}.")
        if arguments.get("format"):
            return _tool_error("Format conversion is only supported for Office documents; this file can't be converted.")
        blob = drive_content(path)
        if blob is None:
            return _tool_error(f"Microsoft Graph returned 404 itemNotFound for {path}.")
        structured = {
            "statusCode": 200,
            "contentType": blob["contentType"],
            "blobName": blob["name"],
            "sizeBytes": len(blob["content"]),
            "base64Content": base64.b64encode(blob["content"]).decode(),
        }
        # Work IQ doesn't duplicate the blob into the text content.
        summary = f"Fetched {blob['name']} ({len(blob['content'])} bytes, {blob['contentType']}); read base64Content from structuredContent."
        return _text_result(summary, structured=structured)

    if name == "ask":
        if _string_argument(arguments, "question") is None:
            return _tool_error("question is required.")
        return _tool_error(
            "The ask tool invokes Microsoft 365 Copilot, which this local simulator doesn't emulate. "
            "Use fetch or call_function against the seeded paths instead."
        )

    if name == "list_agents":
        return _text_result(json.dumps([BUILTIN_AGENT]))

    if name == "search_paths":
        needle = _string_argument(arguments, "filter")
        if needle is None:
            return _tool_error("filter is required.")
        try:
            pattern = re.compile(needle, re.IGNORECASE)
        except re.error:
            pattern = re.compile(re.escape(needle), re.IGNORECASE)
        paths = [{"path": path, "operations": ["fetch"]} for path in PATHS if pattern.search(path)]
        return _json_result({"paths": paths})

    # get_schema
    operation = arguments.get("operationType")
    if operation not in ("fetch", "create", "update"):
        return _tool_error("operationType must be one of fetch, create, or update.")
    path, operation_ids = _string_argument(arguments, "path"), _string_argument(arguments, "operationIds")
    if bool(path) == bool(operation_ids):
        return _tool_error("Provide either path or operationIds, not both.")
    if operation_ids:
        return _tool_error("operationIds lookup isn't simulated; pass a path from search_paths instead.")
    output_format = arguments.get("format") or "jsonschema"
    if output_format not in ("jsonschema", "typescript"):
        return _tool_error("format must be jsonschema or typescript.")
    if output_format == "typescript":
        return _tool_error("TypeScript output isn't simulated; omit format to get JSON Schema.")
    template = _template_for(path)
    if template is None:
        return _tool_error(f"No Microsoft Graph v1.0 path matches '{path}'. Use search_paths to discover paths.")
    if operation != "fetch":
        return _tool_error(f"The {operation} operation isn't available for {template}.")
    # Work IQ returns the schema as TextContent only.
    return _text_result(json.dumps(schema_for(template)))


def mcp_initialize(params: dict) -> dict:
    requested = params.get("protocolVersion")
    version = requested if requested in SUPPORTED_PROTOCOL_VERSIONS else SUPPORTED_PROTOCOL_VERSIONS[0]
    return {
        "protocolVersion": version,
        "capabilities": {"tools": {"listChanged": False}},
        "serverInfo": {"name": "neomind-workiq-sim", "version": "0.2.0"},
        "instructions": "Local Work IQ MCP simulator. Seeded data covers one demo student; mutation tools are denied by default tenant policy.",
    }


def mcp_handle(message: dict) -> dict:
    """Answer one JSON-RPC request (not a notification) with a JSON-RPC response."""
    request_id = message.get("id")
    params = message.get("params", {})
    try:
        if not isinstance(params, dict):
            raise JsonRpcError(-32602, "params must be an object.")
        method = message.get("method")
        if method == "initialize":
            result = mcp_initialize(params)
        elif method == "ping":
            result = {}
        elif method == "tools/list":
            result = {"tools": MCP_TOOLS}
        elif method == "tools/call":
            name, arguments = params.get("name"), params.get("arguments", {})
            if not isinstance(name, str):
                raise JsonRpcError(-32602, "tools/call requires a tool name.")
            if not isinstance(arguments, dict):
                raise JsonRpcError(-32602, "arguments must be an object.")
            result = mcp_call(name, arguments)
        else:
            raise JsonRpcError(-32601, f"Method not found: {method}")
    except JsonRpcError as exc:
        return {"jsonrpc": "2.0", "id": request_id, "error": {"code": exc.code, "message": exc.message}}
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


# --- Entra ID (Microsoft identity platform v2.0) -----------------------------

# Authorization requests waiting for their code to be redeemed, keyed by code.
PENDING_AUTHORIZATIONS: dict[str, dict] = {}


def entra_error(error: str, code: int, description: str) -> dict:
    timestamp = _now().strftime("%Y-%m-%d %H:%M:%SZ")
    trace_id, correlation_id = str(uuid.uuid4()), str(uuid.uuid4())
    return {
        "error": error,
        "error_description": f"AADSTS{code}: {description} Trace ID: {trace_id} Correlation ID: {correlation_id} Timestamp: {timestamp}",
        "error_codes": [code],
        "timestamp": timestamp,
        "trace_id": trace_id,
        "correlation_id": correlation_id,
        "error_uri": f"{BASE_URL}/error?code={code}",
    }


def openid_configuration(tenant: str) -> dict:
    # Multi-tenant aliases publish a templated issuer, like Entra does. Like
    # Entra, there's no code_challenge_methods_supported even though S256 works.
    issuer_tenant = "{tenantid}" if tenant in ("common", "organizations") else TENANT_ID
    base = f"{BASE_URL}/{tenant}"
    return {
        "token_endpoint": f"{base}/oauth2/v2.0/token",
        "token_endpoint_auth_methods_supported": ["client_secret_post", "private_key_jwt", "client_secret_basic"],
        "jwks_uri": f"{base}/discovery/v2.0/keys",
        "response_modes_supported": ["query", "fragment", "form_post"],
        "subject_types_supported": ["pairwise"],
        "id_token_signing_alg_values_supported": ["RS256"],
        "response_types_supported": ["code", "id_token", "code id_token", "id_token token"],
        "scopes_supported": ["openid", "profile", "email", "offline_access"],
        "issuer": f"{BASE_URL}/{issuer_tenant}/v2.0",
        "request_uri_parameter_supported": False,
        "userinfo_endpoint": f"{BASE_URL}/oidc/userinfo",
        "authorization_endpoint": f"{base}/oauth2/v2.0/authorize",
        "claims_supported": ["sub", "iss", "aud", "exp", "iat", "nonce", "preferred_username", "name", "tid", "ver", "oid", "email"],
        "cloud_instance_name": "microsoftonline.com",
        "msgraph_host": "graph.microsoft.com",
    }


def resource_for_scopes(scope: str) -> tuple[str | None, dict | None]:
    """Which resource ("graph" or "workiq") a scope string targets."""
    resources = set()
    for item in scope.split():
        lowered = item.lower()
        if lowered in OIDC_SCOPES:
            continue
        if lowered.startswith(WORKIQ_RESOURCE + "/"):
            resources.add("workiq")
        elif lowered.startswith("https://graph.microsoft.com/") or "/" not in lowered:
            resources.add("graph")  # bare scope names like User.Read are Graph's
        else:
            return None, entra_error("invalid_scope", 70011, f"The provided value for the input parameter 'scope' is not valid. The scope {item} is not valid.")
    if len(resources) > 1:
        return None, entra_error(
            "invalid_scope",
            28000,
            f"Provided value for the input parameter scope is not valid because it contains more than one resource. Scope {scope} is not valid.",
        )
    # Entra defaults to Microsoft Graph when only OIDC scopes are requested.
    return (resources.pop() if resources else "graph"), None


def _unsigned_jwt(claims: dict) -> str:
    # Simulator assumption: Entra signs id_tokens with RS256; this double has no
    # keys, so the token is unsigned ("alg": "none") and the JWKS is empty.
    def encode(part: dict) -> str:
        return base64.urlsafe_b64encode(json.dumps(part).encode()).rstrip(b"=").decode()

    return f"{encode({'typ': 'JWT', 'alg': 'none'})}.{encode(claims)}."


def token_response(form: dict, client_id: str) -> tuple[int, dict]:
    """The /oauth2/v2.0/token endpoint: authorization_code and refresh_token grants."""
    grant_type = form.get("grant_type")
    nonce = None
    if grant_type == "authorization_code":
        if form.get("code") != AUTH_CODE:
            return 400, entra_error("invalid_grant", 70000, "The provided value for the 'code' parameter is not valid.")
        pending = PENDING_AUTHORIZATIONS.get(AUTH_CODE)
        if pending:
            if form.get("redirect_uri") != pending["redirect_uri"]:
                return 400, entra_error(
                    "invalid_grant",
                    500112,
                    "The reply address does not match the reply address provided when requesting the authorization code.",
                )
            if pending["code_challenge"]:
                verifier = form.get("code_verifier", "")
                if pending["code_challenge_method"] == "S256":
                    derived = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
                else:
                    derived = verifier
                if not verifier or derived != pending["code_challenge"]:
                    return 400, entra_error(
                        "invalid_grant",
                        501481,
                        "The Code_Verifier does not match the code_challenge supplied in the authorization request.",
                    )
            nonce = pending["nonce"]
            PENDING_AUTHORIZATIONS.pop(AUTH_CODE)  # codes are single-use
        scope = form.get("scope") or (pending["scope"] if pending else "")
    elif grant_type == "refresh_token":
        if form.get("refresh_token") != REFRESH_TOKEN:
            return 400, entra_error("invalid_grant", 9002313, "Invalid request. Request is malformed or invalid.")
        scope = form.get("scope", "")
    else:
        return 400, entra_error("unsupported_grant_type", 70003, f"The app requested an unsupported grant type '{grant_type}'.")

    resource, error = resource_for_scopes(scope)
    if error:
        return 400, error
    requested = {item.lower() for item in scope.split()}
    oidc = [item for item in ("profile", "openid", "email") if item in requested]
    granted = [WORKIQ_SCOPE] if resource == "workiq" else list(GRAPH_SCOPES)
    body = {
        "token_type": "Bearer",
        "scope": " ".join(granted + oidc),
        "expires_in": 3599,
        "ext_expires_in": 3599,
        "access_token": WORKIQ_TOKEN if resource == "workiq" else GRAPH_TOKEN,
    }
    if "offline_access" in requested or grant_type == "refresh_token":
        body["refresh_token"] = REFRESH_TOKEN
    if "openid" in requested:
        now = int(time.time())
        claims = {
            "aud": client_id,
            "iss": f"{BASE_URL}/{TENANT_ID}/v2.0",
            "iat": now,
            "nbf": now,
            "exp": now + 3600,
            "name": STUDENT["displayName"],
            "oid": STUDENT_ID,
            "preferred_username": UPN,
            "sub": "sim-pairwise-subject-student-001",
            "tid": TENANT_ID,
            "ver": "2.0",
        }
        if nonce:
            claims["nonce"] = nonce
        body["id_token"] = _unsigned_jwt(claims)
    return 200, body


# --- HTTP --------------------------------------------------------------------

GRAPH_CONTENT_TYPE = "application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false;charset=utf-8"
CORS_ORIGIN = "http://localhost:3000"  # the Next.js walkthrough's origin
LOCAL_HOSTNAMES = {"localhost", "127.0.0.1", "::1"}
ENTRA_PATH = re.compile(r"/(?P<tenant>[^/]+)/(?P<endpoint>v2\.0/\.well-known/openid-configuration|discovery/v2\.0/keys|oauth2/v2\.0/authorize|oauth2/v2\.0/token)")


class Handler(BaseHTTPRequestHandler):
    server_version = "NeoMindWorkIQSim/0.2"

    def log_message(self, format: str, *args: object) -> None:
        print("[sim] " + format % args)

    def _send(self, status: int, body: object = None, *, content_type: str = "application/json", headers: dict[str, str] | None = None) -> None:
        if body is None:
            encoded = b""
        elif isinstance(body, bytes):
            encoded = body
        else:
            encoded = json.dumps(body).encode()
        self.send_response(status)
        if encoded:
            self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", CORS_ORIGIN)
        self.send_header("Access-Control-Expose-Headers", "WWW-Authenticate, request-id, client-request-id")
        self.send_header("Content-Length", str(len(encoded)))
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(encoded)

    def _redirect(self, location: str) -> None:
        self.send_response(HTTPStatus.FOUND)
        self.send_header("Location", location)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _body(self) -> bytes:
        return self.rfile.read(int(self.headers.get("Content-Length", "0") or 0))

    def _bearer(self) -> str | None:
        scheme, _, token = (self.headers.get("Authorization") or "").partition(" ")
        return token.strip() if scheme.lower() == "bearer" and token.strip() else None

    def do_OPTIONS(self) -> None:
        self._send(
            HTTPStatus.NO_CONTENT,
            headers={
                "Access-Control-Allow-Headers": "Authorization, Content-Type, Accept, MCP-Protocol-Version, Mcp-Session-Id, client-request-id",
                "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            },
        )

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self._send(200, {"status": "ok"})
        elif parsed.path.startswith("/v1.0/"):
            self._graph_get(parsed)
        elif parsed.path.startswith("/sim-download/"):
            self._download(parsed.path[len("/sim-download/"):])
        elif parsed.path in ("/.well-known/oauth-protected-resource", "/.well-known/oauth-protected-resource/mcp"):
            self._send(
                200,
                {
                    "resource": f"{BASE_URL}/mcp",
                    "authorization_servers": [f"{BASE_URL}/{TENANT_ID}/v2.0"],
                    "scopes_supported": [WORKIQ_SCOPE],
                    "bearer_methods_supported": ["header"],
                },
            )
        elif parsed.path == "/oidc/userinfo":
            if self._bearer() != GRAPH_TOKEN:
                self._send(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Access token validation failure."}})
            else:
                self._send(200, {"sub": "sim-pairwise-subject-student-001", "name": STUDENT["displayName"], "family_name": STUDENT["surname"], "given_name": STUDENT["givenName"], "email": UPN})
        elif parsed.path == "/mcp":
            # No server-initiated SSE stream is offered.
            self._send(HTTPStatus.METHOD_NOT_ALLOWED, headers={"Allow": "POST"})
        elif match := ENTRA_PATH.fullmatch(parsed.path):
            self._entra(match, parsed)
        else:
            self._send(404, {"error": "not_found"})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/mcp":
            self._mcp()
        elif parsed.path.startswith("/v1.0/"):
            # The simulated user only holds read scopes, so writes fail like
            # they would in Graph for an under-privileged token.
            self._send(403, graph_error(403, "Authorization_RequestDenied", "Insufficient privileges to complete the operation.")[1], content_type=GRAPH_CONTENT_TYPE)
        elif (match := ENTRA_PATH.fullmatch(parsed.path)) and match["endpoint"] == "oauth2/v2.0/token":
            self._entra(match, parsed)
        else:
            self._send(404, {"error": "not_found"})

    def do_DELETE(self) -> None:
        if urlparse(self.path).path == "/mcp":
            self._send(HTTPStatus.METHOD_NOT_ALLOWED, headers={"Allow": "POST"})
        else:
            self._send(404, {"error": "not_found"})

    # Graph ------------------------------------------------------------------

    def _graph_get(self, parsed) -> None:
        request_id = self.headers.get("client-request-id") or str(uuid.uuid4())
        headers = {"request-id": request_id, "client-request-id": request_id}
        token = self._bearer()
        if token != GRAPH_TOKEN:
            if token is None:
                message = "Access token is empty."
            elif token == WORKIQ_TOKEN:
                message = "Access token validation failure. Invalid audience."
            else:
                message = "IDX14100: JWT is not well formed, there are no dots (.)."
            headers["WWW-Authenticate"] = f'Bearer realm="", authorization_uri="{BASE_URL}/common/oauth2/v2.0/authorize", client_id="00000003-0000-0000-c000-000000000000"'
            self._send(401, graph_error(401, "InvalidAuthenticationToken", message)[1], content_type=GRAPH_CONTENT_TYPE, headers=headers)
            return
        relative = parsed.path[len("/v1.0"):]
        if drive_content(relative) is not None:
            # Graph answers /content with a redirect to a pre-authenticated URL.
            item_id = unquote(relative).split("/")[4]
            self._redirect(f"{BASE_URL}/sim-download/{item_id}")
            return
        status, data = graph_fetch(relative + (("?" + parsed.query) if parsed.query else ""))
        self._send(status, data, content_type=GRAPH_CONTENT_TYPE, headers=headers)

    def _download(self, item_id: str) -> None:
        blob = drive_content(f"/me/drive/items/{item_id}/content")
        if blob is None:
            self._send(404, {"error": "not_found"})
        else:
            self._send(200, blob["content"], content_type=blob["contentType"])

    # MCP --------------------------------------------------------------------

    def _mcp(self) -> None:
        origin = self.headers.get("Origin")
        if origin and urlparse(origin).hostname not in LOCAL_HOSTNAMES:
            # DNS-rebinding protection required by the Streamable HTTP transport.
            self._send(403, {"jsonrpc": "2.0", "id": None, "error": {"code": -32600, "message": "Origin not allowed."}})
            return
        token = self._bearer()
        if token != WORKIQ_TOKEN:
            challenge = f'Bearer resource_metadata="{BASE_URL}/.well-known/oauth-protected-resource/mcp", scope="{WORKIQ_SCOPE}"'
            if token:
                challenge += ', error="invalid_token"'
            self._send(401, {"jsonrpc": "2.0", "id": None, "error": {"code": -32001, "message": "Unauthorized"}}, headers={"WWW-Authenticate": challenge})
            return
        version = self.headers.get("MCP-Protocol-Version")
        if version and version not in SUPPORTED_PROTOCOL_VERSIONS:
            self._send(400, {"jsonrpc": "2.0", "id": None, "error": {"code": -32600, "message": f"Unsupported MCP-Protocol-Version: {version}"}})
            return
        try:
            message = json.loads(self._body() or b"null")
        except (ValueError, UnicodeDecodeError):
            self._send(400, {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error"}})
            return
        # JSON-RPC batches were removed from MCP in 2025-06-18.
        if not isinstance(message, dict) or message.get("jsonrpc") != "2.0":
            self._send(400, {"jsonrpc": "2.0", "id": None, "error": {"code": -32600, "message": "Invalid Request"}})
            return
        if "method" not in message or "id" not in message:
            # Notifications and client responses are acknowledged without a body.
            self._send(HTTPStatus.ACCEPTED)
            return
        self._send(200, mcp_handle(message))

    # Entra ------------------------------------------------------------------

    def _entra(self, match: re.Match, parsed) -> None:
        tenant = match["tenant"]
        if tenant not in TENANT_ALIASES:
            self._send(400, entra_error("invalid_request", 90002, f"Tenant '{tenant}' not found. Check to make sure you have the correct tenant ID and are signing into the correct cloud."))
            return
        endpoint = match["endpoint"]
        if endpoint.endswith("openid-configuration"):
            self._send(200, openid_configuration(tenant))
        elif endpoint.endswith("keys"):
            self._send(200, {"keys": []})
        elif endpoint.endswith("authorize"):
            self._authorize({key: values[-1] for key, values in parse_qs(parsed.query).items()})
        else:
            form = {key: values[-1] for key, values in parse_qs(self._body().decode()).items()}
            client_id = form.get("client_id")
            if not client_id and self.headers.get("Authorization", "").lower().startswith("basic "):
                client_id = base64.b64decode(self.headers["Authorization"][6:]).decode().partition(":")[0]
            if not client_id:
                self._send(400, entra_error("invalid_request", 900144, "The request body must contain the following parameter: 'client_id'."))
                return
            status, body = token_response(form, client_id)
            self._send(status, body, headers={"Cache-Control": "no-store", "Pragma": "no-cache"})

    def _authorize(self, query: dict) -> None:
        for required in ("client_id", "redirect_uri"):
            if not query.get(required):
                # Without a trusted redirect URI, Entra shows an error page instead of redirecting.
                self._send(400, entra_error("invalid_request", 900144, f"The request body must contain the following parameter: '{required}'."))
                return
        redirect_uri, state = query["redirect_uri"], query.get("state")
        mode = query.get("response_mode", "query")

        def respond(params: dict) -> None:
            if state is not None:
                params["state"] = state
            if mode == "form_post":
                fields = "".join(f'<input type="hidden" name="{escape(key)}" value="{escape(value)}">' for key, value in params.items())
                page = f'<html><body onload="document.forms[0].submit()"><form method="post" action="{escape(redirect_uri)}">{fields}</form></body></html>'
                self._send(200, page.encode(), content_type="text/html; charset=utf-8")
            else:
                separator = "#" if mode == "fragment" else ("&" if "?" in redirect_uri else "?")
                self._redirect(redirect_uri + separator + urlencode(params))

        def fail(error: str, code: int, description: str) -> None:
            respond({"error": error, "error_description": f"AADSTS{code}: {description}"})

        if query.get("response_type") != "code":
            fail("unsupported_response_type", 700054, f"response_type '{query.get('response_type', '')}' is not enabled for the application.")
            return
        if not query.get("scope", "").strip():
            fail("invalid_request", 900144, "The request body must contain the following parameter: 'scope'.")
            return
        challenge = query.get("code_challenge")
        method = query.get("code_challenge_method", "plain")
        if challenge and method not in ("S256", "plain"):
            fail("invalid_request", 50148, f"The code_challenge_method '{method}' is not supported.")
            return
        PENDING_AUTHORIZATIONS[AUTH_CODE] = {
            "redirect_uri": redirect_uri,
            "scope": query["scope"],
            "code_challenge": challenge,
            "code_challenge_method": method,
            "nonce": query.get("nonce"),
        }
        respond({"code": AUTH_CODE})


if __name__ == "__main__":
    print(f"NeoMind Work IQ simulation listening at {BASE_URL}")
    try:
        ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        pass
