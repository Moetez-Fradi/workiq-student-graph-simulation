import base64
import hashlib
import json
import threading
import unittest
from http.client import HTTPConnection
from http.server import ThreadingHTTPServer
from urllib.parse import parse_qs, urlencode, urlparse

import server


class GraphFetchTests(unittest.TestCase):
    def test_student_classes_are_graph_shaped(self):
        status, body = server.graph_fetch(f"/education/users/{server.STUDENT_ID}/classes")
        self.assertEqual(status, 200)
        self.assertTrue(body["@odata.context"].startswith("https://graph.microsoft.com/v1.0/$metadata#"))
        self.assertEqual(body["value"][0]["id"], server.CLASS_ID)

    def test_education_me_aliases_the_signed_in_student(self):
        self.assertEqual(server.graph_fetch("/education/me/classes"), server.graph_fetch(f"/education/users/{server.STUDENT_ID}/classes"))
        status, body = server.graph_fetch("/education/me")
        self.assertEqual(status, 200)
        self.assertEqual(body["primaryRole"], "student")
        self.assertIn("student", body)

    def test_class_course_maps_to_the_neomind_curriculum(self):
        _, body = server.graph_fetch("/education/me/classes")
        self.assertEqual(body["value"][0]["course"]["externalId"], "primary_math")
        _, schools = server.graph_fetch("/education/schools")
        self.assertNotIn("neomindCurriculumId", schools["value"][0])

    def test_members_are_education_users(self):
        _, body = server.graph_fetch(f"/education/classes/{server.CLASS_ID}/members")
        self.assertEqual(body["@odata.context"], server.GRAPH_CONTEXT + "Collection(microsoft.graph.educationUser)")
        self.assertEqual(body["value"][0]["primaryRole"], "student")

    def test_students_see_published_assignments_as_assigned(self):
        _, body = server.graph_fetch(f"/education/classes/{server.CLASS_ID}/assignments")
        self.assertEqual(body["value"][0]["status"], "assigned")
        self.assertEqual(body["value"][0]["classId"], server.CLASS_ID)

    def test_paths_are_case_insensitive(self):
        status, _ = server.graph_fetch("/Education/Me/Classes")
        self.assertEqual(status, 200)

    def test_unknown_id_is_a_graph_not_found(self):
        status, body = server.graph_fetch("/education/users/00000000-0000-0000-0000-000000000000")
        self.assertEqual(status, 404)
        self.assertEqual(body["error"]["code"], "Request_ResourceNotFound")
        self.assertIn("request-id", body["error"]["innerError"])

    def test_unknown_segment_is_a_bad_request(self):
        status, body = server.graph_fetch("/education/unicorns")
        self.assertEqual(status, 400)
        self.assertEqual(body["error"]["message"], "Resource not found for the segment 'unicorns'.")

    def test_calendar_view_requires_and_applies_a_time_window(self):
        status, body = server.graph_fetch("/me/calendarView")
        self.assertEqual((status, body["error"]["code"]), (400, "ErrorInvalidParameter"))
        _, inside = server.graph_fetch("/me/calendarView?startDateTime=2026-09-20T00:00:00Z&endDateTime=2026-09-21T00:00:00Z")
        _, outside = server.graph_fetch("/me/calendarview?startdatetime=2026-10-01T00:00:00Z&enddatetime=2026-10-02T00:00:00Z")
        self.assertEqual(len(inside["value"]), 1)
        self.assertEqual(outside["value"], [])

    def test_drive_search_matches_any_query(self):
        _, hit = server.graph_fetch("/me/drive/root/search(q='fractions')")
        _, encoded_hit = server.graph_fetch("/me/drive/root/search(q=%27primary_math%27)")
        _, miss = server.graph_fetch("/me/drive/root/search(q='chemistry')")
        self.assertEqual(hit["value"][0]["file"]["mimeType"], "text/plain")
        self.assertEqual(len(encoded_hit["value"]), 1)
        self.assertEqual(miss["value"], [])

    def test_select_and_top_query_options(self):
        _, body = server.graph_fetch("/me?$select=displayName,mail")
        self.assertEqual(set(body) - {"@odata.context"}, {"displayName", "mail"})
        _, body = server.graph_fetch("/education/schools?$top=0")
        self.assertEqual(body["value"], [])


class WorkIQToolTests(unittest.TestCase):
    def test_fetch_returns_one_result_per_path_in_order(self):
        result = server.mcp_call("fetch", {"entityUrls": ["/me", "/me/calendarView"]})
        results = result["structuredContent"]["results"]
        self.assertEqual([entry["statusCode"] for entry in results], [200, 400])
        self.assertEqual(results[0]["data"]["id"], server.STUDENT_ID)
        self.assertEqual(json.loads(result["content"][0]["text"]), result["structuredContent"])

    def test_default_policy_blocks_paths_outside_me_users_and_sites(self):
        results = server.mcp_call("fetch", {"entityUrls": ["/education/me/classes", "/servicePrincipals/x/authentication/y"]})
        statuses = [entry["statusCode"] for entry in results["structuredContent"]["results"]]
        self.assertEqual(statuses, [403, 403])
        self.assertIn("policy", results["structuredContent"]["results"][0]["data"]["error"]["message"])

    def test_default_policy_blocks_skip_and_caps_top(self):
        skipped = server.mcp_call("fetch", {"entityUrls": ["/me/drive/root/search(q='fractions')?$skip=1"]})
        self.assertEqual(skipped["structuredContent"]["results"][0]["statusCode"], 403)
        capped = server.mcp_call("fetch", {"entityUrls": ["/me/drive/root/search(q='fractions')?$top=500"]})
        self.assertEqual(capped["structuredContent"]["results"][0]["statusCode"], 200)

    def test_call_function_reads_calendar_view(self):
        result = server.mcp_call("call_function", {"functionUrl": "/me/calendarView?startdatetime=2026-09-20T00:00:00Z&enddatetime=2026-09-21T00:00:00Z"})
        self.assertEqual(result["structuredContent"]["statusCode"], 200)
        self.assertEqual(len(result["structuredContent"]["data"]["value"]), 1)

    def test_fetch_blob_returns_base64_in_structured_content_only(self):
        result = server.mcp_call("fetch_blob", {"path": f"/me/drive/items/{server.DRIVE_ITEM['id']}/content"})
        blob = result["structuredContent"]
        self.assertEqual(base64.b64decode(blob["base64Content"]), server.NOTES_CONTENT)
        self.assertEqual(blob["sizeBytes"], len(server.NOTES_CONTENT))
        self.assertNotIn(blob["base64Content"], result["content"][0]["text"])

    def test_search_paths_accepts_regex_and_uses_graph_templates(self):
        result = server.mcp_call("search_paths", {"filter": ".*assignments$"})
        paths = result["structuredContent"]["paths"]
        self.assertEqual(paths, [{"path": "/education/classes/{educationClass-id}/assignments", "operations": ["fetch"]}])

    def test_get_schema_returns_a_collection_schema_as_text(self):
        result = server.mcp_call("get_schema", {"path": "/education/me/classes", "operationType": "fetch"})
        self.assertNotIn("structuredContent", result)
        schema = json.loads(result["content"][0]["text"])
        self.assertEqual(schema["title"], "microsoft.graph.educationClassCollectionResponse")
        self.assertIn("course", schema["$defs"]["microsoft.graph.educationClass"]["properties"])

    def test_get_schema_requires_exactly_one_of_path_or_operation_ids(self):
        self.assertTrue(server.mcp_call("get_schema", {"operationType": "fetch"})["isError"])

    def test_mutation_tools_are_denied_by_policy(self):
        for name in server.MUTATION_TOOLS:
            result = server.mcp_call(name, {})
            self.assertTrue(result["isError"])
            self.assertIn("policy", result["content"][0]["text"])

    def test_unknown_tool_is_a_protocol_error(self):
        with self.assertRaises(server.JsonRpcError) as caught:
            server.mcp_call("send_mail", {})
        self.assertEqual(caught.exception.code, -32602)

    def test_tools_list_matches_the_work_iq_surface(self):
        names = {tool["name"] for tool in server.MCP_TOOLS}
        self.assertEqual(
            names,
            {"fetch", "fetch_blob", "create_entity", "update_entity", "delete_entity", "do_action", "call_function", "ask", "list_agents", "get_schema", "search_paths"},
        )


class HttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = ThreadingHTTPServer(("127.0.0.1", 0), server.Handler)
        cls.port = cls.httpd.server_address[1]
        threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def request(self, method, path, body=None, headers=None):
        connection = HTTPConnection("127.0.0.1", self.port)
        if isinstance(body, dict):
            body = json.dumps(body)
        connection.request(method, path, body=body, headers=headers or {})
        response = connection.getresponse()
        raw = response.read()
        connection.close()
        payload = json.loads(raw) if raw and "json" in (response.getheader("Content-Type") or "") else raw
        return response.status, dict(response.getheaders()), payload

    def mcp(self, message, token=server.WORKIQ_TOKEN, **headers):
        headers = {"Content-Type": "application/json", "Accept": "application/json, text/event-stream", **headers}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        return self.request("POST", "/mcp", message, headers)

    def token(self, **form):
        return self.request("POST", f"/{server.TENANT_ID}/oauth2/v2.0/token", urlencode(form), {"Content-Type": "application/x-www-form-urlencoded"})

    def test_initialize_negotiates_the_protocol_version(self):
        _, _, known = self.mcp({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18"}})
        _, _, unknown = self.mcp({"jsonrpc": "2.0", "id": 2, "method": "initialize", "params": {"protocolVersion": "1999-01-01"}})
        self.assertEqual(known["result"]["protocolVersion"], "2025-06-18")
        self.assertEqual(unknown["result"]["protocolVersion"], server.SUPPORTED_PROTOCOL_VERSIONS[0])

    def test_notifications_are_accepted_without_a_body(self):
        status, _, body = self.mcp({"jsonrpc": "2.0", "method": "notifications/initialized"})
        self.assertEqual((status, body), (202, b""))

    def test_unknown_tool_is_a_jsonrpc_invalid_params_error(self):
        _, _, body = self.mcp({"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "nope", "arguments": {}}})
        self.assertEqual(body["error"]["code"], -32602)

    def test_malformed_json_and_batches_are_rejected(self):
        status, _, body = self.mcp("{not json")
        self.assertEqual((status, body["error"]["code"]), (400, -32700))
        status, _, body = self.request("POST", "/mcp", json.dumps([{"jsonrpc": "2.0", "id": 1, "method": "ping"}]), {"Authorization": f"Bearer {server.WORKIQ_TOKEN}"})
        self.assertEqual((status, body["error"]["code"]), (400, -32600))

    def test_mcp_rejects_graph_tokens_with_a_discovery_challenge(self):
        status, headers, _ = self.mcp({"jsonrpc": "2.0", "id": 4, "method": "ping"}, token=server.GRAPH_TOKEN)
        self.assertEqual(status, 401)
        self.assertIn('resource_metadata="', headers["WWW-Authenticate"])
        self.assertIn('error="invalid_token"', headers["WWW-Authenticate"])
        _, _, metadata = self.request("GET", "/.well-known/oauth-protected-resource/mcp")
        self.assertEqual(metadata["scopes_supported"], [server.WORKIQ_SCOPE])

    def test_mcp_rejects_foreign_origins_and_unsupported_versions(self):
        status, _, _ = self.mcp({"jsonrpc": "2.0", "id": 5, "method": "ping"}, Origin="https://evil.example")
        self.assertEqual(status, 403)
        status, _, _ = self.mcp({"jsonrpc": "2.0", "id": 6, "method": "ping"}, **{"MCP-Protocol-Version": "1999-01-01"})
        self.assertEqual(status, 400)

    def test_mcp_get_is_method_not_allowed(self):
        status, _, _ = self.request("GET", "/mcp", headers={"Authorization": f"Bearer {server.WORKIQ_TOKEN}"})
        self.assertEqual(status, 405)

    def test_graph_rejects_work_iq_tokens(self):
        status, _, body = self.request("GET", "/v1.0/me", headers={"Authorization": f"Bearer {server.WORKIQ_TOKEN}"})
        self.assertEqual(status, 401)
        self.assertEqual(body["error"]["message"], "Access token validation failure. Invalid audience.")
        status, headers, body = self.request("GET", "/v1.0/me", headers={"Authorization": f"Bearer {server.GRAPH_TOKEN}"})
        self.assertEqual((status, body["id"]), (200, server.STUDENT_ID))
        self.assertTrue(headers["Content-Type"].startswith("application/json;odata.metadata=minimal"))

    def test_graph_drive_content_redirects_to_a_download_url(self):
        status, headers, _ = self.request("GET", f"/v1.0/me/drive/items/{server.DRIVE_ITEM['id']}/content", headers={"Authorization": f"Bearer {server.GRAPH_TOKEN}"})
        self.assertEqual(status, 302)
        status, _, body = self.request("GET", urlparse(headers["Location"]).path)
        self.assertEqual((status, body), (200, server.NOTES_CONTENT))

    def test_openid_configuration_is_entra_shaped(self):
        _, _, config = self.request("GET", "/common/v2.0/.well-known/openid-configuration")
        self.assertTrue(config["issuer"].endswith("/{tenantid}/v2.0"))
        self.assertTrue(config["token_endpoint"].endswith("/common/oauth2/v2.0/token"))
        status, _, body = self.request("GET", "/unknown-tenant/v2.0/.well-known/openid-configuration")
        self.assertEqual((status, body["error_codes"]), (400, [90002]))

    def test_authorization_code_flow_with_pkce_issues_a_work_iq_token(self):
        verifier = "a" * 43
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
        query = urlencode({
            "client_id": "neomind-demo",
            "response_type": "code",
            "redirect_uri": "http://127.0.0.1:3000/callback",
            "scope": f"{server.WORKIQ_SCOPE} openid offline_access",
            "state": "demo",
            "code_challenge": challenge,
            "code_challenge_method": "S256",
            "nonce": "n-1",
        })
        status, headers, _ = self.request("GET", f"/common/oauth2/v2.0/authorize?{query}")
        self.assertEqual(status, 302)
        params = parse_qs(urlparse(headers["Location"]).query)
        self.assertEqual((params["code"], params["state"]), ([server.AUTH_CODE], ["demo"]))

        status, _, body = self.token(grant_type="authorization_code", client_id="neomind-demo", code=server.AUTH_CODE, redirect_uri="http://127.0.0.1:3000/callback", code_verifier=verifier)
        self.assertEqual(status, 200)
        self.assertEqual(body["access_token"], server.WORKIQ_TOKEN)
        self.assertEqual(body["refresh_token"], server.REFRESH_TOKEN)
        claims = json.loads(base64.urlsafe_b64decode(body["id_token"].split(".")[1] + "=="))
        self.assertEqual((claims["aud"], claims["nonce"], claims["oid"]), ("neomind-demo", "n-1", server.STUDENT_ID))

        # The refresh token redeems a token for the other resource, like Entra's.
        status, _, body = self.token(grant_type="refresh_token", client_id="neomind-demo", refresh_token=server.REFRESH_TOKEN, scope="EduRoster.ReadBasic")
        self.assertEqual((status, body["access_token"]), (200, server.GRAPH_TOKEN))

    def test_pkce_mismatch_is_rejected(self):
        query = urlencode({"client_id": "c", "response_type": "code", "redirect_uri": "http://localhost/cb", "scope": "User.Read", "code_challenge": "x" * 43, "code_challenge_method": "S256"})
        self.request("GET", f"/common/oauth2/v2.0/authorize?{query}")
        status, _, body = self.token(grant_type="authorization_code", client_id="c", code=server.AUTH_CODE, redirect_uri="http://localhost/cb", code_verifier="wrong")
        self.assertEqual((status, body["error"], body["error_codes"]), (400, "invalid_grant", [501481]))
        server.PENDING_AUTHORIZATIONS.clear()

    def test_one_token_cannot_span_two_resources(self):
        status, _, body = self.token(grant_type="authorization_code", client_id="c", code=server.AUTH_CODE, scope=f"User.Read {server.WORKIQ_SCOPE}")
        self.assertEqual((status, body["error"], body["error_codes"]), (400, "invalid_scope", [28000]))

    def test_authorize_errors_redirect_back_with_aadsts_codes(self):
        query = urlencode({"client_id": "c", "response_type": "token", "redirect_uri": "http://localhost/cb", "scope": "User.Read", "state": "s"})
        status, headers, _ = self.request("GET", f"/common/oauth2/v2.0/authorize?{query}")
        params = parse_qs(urlparse(headers["Location"]).query)
        self.assertEqual((status, params["error"], params["state"]), (302, ["unsupported_response_type"], ["s"]))
        self.assertTrue(params["error_description"][0].startswith("AADSTS700054"))


if __name__ == "__main__":
    unittest.main()
