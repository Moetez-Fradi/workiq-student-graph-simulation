import unittest

import server


class GraphFetchTests(unittest.TestCase):
    def test_student_classes_are_graph_shaped(self):
        status, body = server.graph_fetch("/education/users/student-001/classes")
        self.assertEqual(status, 200)
        self.assertEqual(body["value"][0]["id"], "primary-math-2026")

    def test_seeded_school_resolves_a_neomind_curriculum(self):
        status, body = server.graph_fetch("/education/schools")
        self.assertEqual(status, 200)
        self.assertEqual(body["value"][0]["neomindCurriculumId"], "primary_math")

    def test_unknown_path_is_a_graph_not_found(self):
        status, body = server.graph_fetch("/education/users/unknown")
        self.assertEqual(status, 404)
        self.assertEqual(body["error"]["code"], "Request_ResourceNotFound")


class McpTests(unittest.TestCase):
    def test_fetch_returns_one_result_per_path(self):
        result = server.mcp_call("fetch", {"entityUrls": ["/me", "/education/users/student-001/classes"]})
        results = result["structuredContent"]["results"]
        self.assertEqual([entry["statusCode"] for entry in results], [200, 200])
        self.assertEqual(results[0]["data"]["id"], "student-001")

    def test_search_paths_discovers_education_paths(self):
        result = server.mcp_call("search_paths", {"filter": "assignments"})
        paths = result["structuredContent"]["paths"]
        self.assertEqual(paths[0]["path"], "/education/classes/{id}/assignments")

    def test_mutation_tool_is_denied(self):
        result = server.mcp_call("create_entity", {})
        self.assertTrue(result["isError"])


if __name__ == "__main__":
    unittest.main()
