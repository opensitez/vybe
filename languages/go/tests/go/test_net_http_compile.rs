//! net/url and net/http: Parse, query escape, Values encoding, Client Get, status constants.

go_compile_cases! {
    // net/url — Parse
    url_parse_https_absolute => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://example.com/path\"); _ = u.Scheme }",
    url_parse_http_with_port => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"http://localhost:8080/api\"); _ = u.Host }",
    url_parse_relative_path_only => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"/files/report.pdf\"); _ = u.Path }",
    url_parse_query_string_fragment => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/search?q=go&lang=en#results\"); _ = u.RawQuery; _ = u.Fragment }",
    url_parse_userinfo_in_authority => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://user:pass@api.example.com/v1\"); _ = u.User }",
    url_parse_request_uri_no_fragment => "package main; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"/index.html?tab=1\"); _ = u.Path }",
    url_parse_result_string_method => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://a.b/c\"); _ = u.String() }",

    // net/url — Escape / Unescape
    url_query_escape_spaces_and_specials => "package main; import \"net/url\"; func main() { _ = url.QueryEscape(\"a b&c=d\") }",
    url_query_unescape_percent_plus => "package main; import \"net/url\"; func main() { _ = url.QueryUnescape(\"hello+world%21\") }",
    url_path_escape_preserves_slashes => "package main; import \"net/url\"; func main() { _ = url.PathEscape(\"dir/file name.txt\") }",
    url_path_unescape_encoded_segments => "package main; import \"net/url\"; func main() { _ = url.PathUnescape(\"a%2Fb%20c\") }",

    // net/url — Values (query building)
    url_values_encode_single_pair => "package main; import \"net/url\"; func main() { v := url.Values{}; v.Set(\"q\", \"golang\"); _ = v.Encode() }",
    url_values_add_duplicate_keys => "package main; import \"net/url\"; func main() { v := url.Values{}; v.Add(\"tag\", \"a\"); v.Add(\"tag\", \"b\"); _ = v.Encode() }",
    url_values_get_missing_returns_empty => "package main; import \"net/url\"; func main() { v := url.Values{}; _ = v.Get(\"missing\") == \"\" }",
    url_values_del_removes_key => "package main; import \"net/url\"; func main() { v := url.Values{\"k\": []string{\"1\"}}; v.Del(\"k\"); _ = len(v) }",
    url_parse_then_set_query_via_values => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/\"); q := url.Values{}; q.Set(\"page\", \"2\"); u.RawQuery = q.Encode(); _ = u.String() }",

    // net/http — Client and Get
    http_client_zero_value_used => "package main; import \"net/http\"; func main() { var c http.Client; _, _ = c.Get(\"https://example.com\") }",
    http_get_package_level_function => "package main; import \"net/http\"; func main() { _, _ = http.Get(\"https://example.com/health\") }",
    http_client_get_assign_response_fields => "package main; import \"net/http\"; func main() { resp, err := http.DefaultClient.Get(\"https://api.test/items\"); if err == nil { _ = resp.StatusCode; _ = resp.Header; _ = resp.Body } }",
    http_client_with_timeout_field => "package main; import \"net/http\"; import \"time\"; func main() { c := http.Client{Timeout: 5 * time.Second}; _, _ = c.Get(\"https://slow.example/status\") }",
    http_new_request_get_no_body => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://example.com\", nil); _ = req.URL }",
    http_new_request_post_with_reader => "package main; import \"net/http\"; import \"strings\"; func main() { body := strings.NewReader(\"payload\"); _, _ = http.NewRequest(http.MethodPost, \"https://example.com/submit\", body) }",
    http_post_with_content_type_header => "package main; import \"net/http\"; import \"strings\"; func main() { _, _ = http.Post(\"https://example.com/form\", \"application/x-www-form-urlencoded\", strings.NewReader(\"a=1\")) }",

    // net/http — status and method constants
    http_status_ok_constant_compare => "package main; import \"net/http\"; func main() { code := http.StatusOK; _ = code == 200 }",
    http_status_not_found_constant => "package main; import \"net/http\"; func main() { _ = http.StatusNotFound == 404 }",
    http_status_internal_server_error_constant => "package main; import \"net/http\"; func main() { _ = http.StatusInternalServerError == 500 }",
    http_status_text_for_ok => "package main; import \"net/http\"; func main() { _ = http.StatusText(http.StatusOK) }",
    http_method_get_constant_in_request => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodGet, \"https://example.com\", nil) }",
    http_method_post_constant_in_request => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodPost, \"https://example.com\", nil) }",
    http_header_set_and_get_on_request => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://example.com\", nil); req.Header.Set(\"Accept\", \"application/json\"); _ = req.Header.Get(\"Accept\") }",
}
