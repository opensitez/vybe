//! net/url extended: ParseRequestURI, Query mutations, JoinPath, ResolveReference, User credentials.

go_run_cases! {
    url_parse_request_uri_path_and_query => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"/api/v2/items?page=3\"); fmt.Println(u.Path); fmt.Println(u.Query().Get(\"page\")) }",
        vec!["/api/v2/items", "3"]
    ),
    url_parse_request_uri_absolute_http => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"http://example.com/health\"); fmt.Println(u.Scheme); fmt.Println(u.Host) }",
        vec!["http", "example.com"]
    ),
    url_query_get_from_parsed_url => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/search?q=vybe&lang=go\"); fmt.Println(u.Query().Get(\"q\")); fmt.Println(u.Query().Get(\"lang\")) }",
        vec!["vybe", "go"]
    ),
    url_query_get_missing_key_empty => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/?a=1\"); fmt.Println(u.Query().Get(\"missing\") == \"\") }",
        vec!["true"]
    ),
    url_query_set_replaces_single_value => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/\"); q := u.Query(); q.Set(\"token\", \"abc\"); u.RawQuery = q.Encode(); fmt.Println(u.Query().Get(\"token\")) }",
        vec!["abc"]
    ),
    url_query_add_appends_duplicate_key => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/\"); q := u.Query(); q.Add(\"id\", \"1\"); q.Add(\"id\", \"2\"); fmt.Println(len(q[\"id\"])) }",
        vec!["2"]
    ),
    url_query_del_removes_all_for_key => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/?x=1&x=2\"); q := u.Query(); q.Del(\"x\"); fmt.Println(len(q)) }",
        vec!["0"]
    ),
    url_query_encode_sorted_keys => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { q := url.Values{}; q.Set(\"b\", \"2\"); q.Set(\"a\", \"1\"); fmt.Println(q.Encode()) }",
        vec!["a=1&b=2"]
    ),
    url_query_encode_spaces_as_plus => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { q := url.Values{}; q.Set(\"msg\", \"hello world\"); fmt.Println(q.Encode()) }",
        vec!["msg=hello+world"]
    ),
    url_path_escape_preserves_slashes => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { fmt.Println(url.PathEscape(\"dir/sub/file name\")) }",
        vec!["dir/sub/file%20name"]
    ),
    url_path_unescape_percent_encoding => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { s, _ := url.PathUnescape(\"a%2Fb%20c\"); fmt.Println(s) }",
        vec!["a/b c"]
    ),
    url_join_path_two_segments => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { fmt.Println(url.JoinPath(\"/a\", \"b\", \"c\")) }",
        vec!["/a/b/c"]
    ),
    url_join_path_elides_dot_dot => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { fmt.Println(url.JoinPath(\"/a/b\", \"../c\")) }",
        vec!["/a/c"]
    ),
    url_resolve_reference_relative_path => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { base, _ := url.Parse(\"https://example.com/a/b\"); ref, _ := url.Parse(\"c\"); fmt.Println(base.ResolveReference(ref).String()) }",
        vec!["https://example.com/a/c"]
    ),
    url_resolve_reference_absolute_override => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { base, _ := url.Parse(\"https://example.com/old\"); ref, _ := url.Parse(\"https://other/new\"); fmt.Println(base.ResolveReference(ref).String()) }",
        vec!["https://other/new"]
    ),
    url_user_username_without_password => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://alice@api.example/v1\"); fmt.Println(u.User.Username()) }",
        vec!["alice"]
    ),
    url_user_password_present => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://bob:secret@host/\"); pw, ok := u.User.Password(); fmt.Println(pw); fmt.Println(ok) }",
        vec!["secret", "true"]
    ),
    url_user_string_roundtrip => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://user:pass@host/\"); fmt.Println(u.User.String()) }",
        vec!["user:pass"]
    ),
    url_parse_request_uri_rejects_fragment => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { _, err := url.ParseRequestURI(\"/page#frag\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    url_raw_query_roundtrip_via_encode => (
        "package main; import \"fmt\"; import \"net/url\"; func main() { u, _ := url.Parse(\"https://h/?k=v\"); q := u.Query(); u.RawQuery = q.Encode(); fmt.Println(u.RawQuery) }",
        vec!["k=v"]
    ),
}

go_compile_cases! {
    url_parse_request_uri_with_percent_encoded => "package main; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"/files/hello%20world.txt\"); _ = u.Path }",
    url_parse_request_uri_https_with_port => "package main; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"https://localhost:8443/status\"); _ = u.Host }",
    url_query_has_multiple_values_for_key => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://h/?tag=a&tag=b\"); _ = u.Query()[\"tag\"] }",
    url_query_set_on_existing_overwrites => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://h/?x=old\"); q := u.Query(); q.Set(\"x\", \"new\"); u.RawQuery = q.Encode(); _ = u.Query().Get(\"x\") }",
    url_query_add_to_existing_key => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://h/?k=1\"); q := u.Query(); q.Add(\"k\", \"2\"); _ = len(q[\"k\"]) }",
    url_query_del_one_of_many_keys => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://h/?a=1&b=2\"); q := u.Query(); q.Del(\"a\"); _ = q.Get(\"b\") }",
    url_query_encode_empty_values => "package main; import \"net/url\"; func main() { q := url.Values{\"empty\": []string{\"\"}}; _ = q.Encode() }",
    url_query_encode_special_characters => "package main; import \"net/url\"; func main() { q := url.Values{}; q.Set(\"q\", \"a&b=c\"); _ = q.Encode() }",
    url_values_get_first_of_duplicates => "package main; import \"net/url\"; func main() { v := url.Values{\"k\": []string{\"first\", \"second\"}}; _ = v.Get(\"k\") }",
    url_path_escape_leading_slash => "package main; import \"net/url\"; func main() { _ = url.PathEscape(\"/root/etc\") }",
    url_path_unescape_plus_not_space => "package main; import \"net/url\"; func main() { _, _ = url.PathUnescape(\"100%25\") }",
    url_join_path_single_base => "package main; import \"net/url\"; func main() { _, _ = url.JoinPath(\"/only\") }",
    url_join_path_empty_elements_skipped => "package main; import \"net/url\"; func main() { _, _ = url.JoinPath(\"a\", \"\", \"b\") }",
    url_join_path_dot_segment => "package main; import \"net/url\"; func main() { _, _ = url.JoinPath(\"/x\", \".\", \"y\") }",
    url_resolve_reference_query_merge => "package main; import \"net/url\"; func main() { base, _ := url.Parse(\"https://ex.com/a?x=1\"); ref, _ := url.Parse(\"?y=2\"); _ = base.ResolveReference(ref).RawQuery }",
    url_resolve_reference_fragment_from_ref => "package main; import \"net/url\"; func main() { base, _ := url.Parse(\"https://ex.com/page\"); ref, _ := url.Parse(\"#section\"); _ = base.ResolveReference(ref).Fragment }",
    url_resolve_reference_parent_directory => "package main; import \"net/url\"; func main() { base, _ := url.Parse(\"https://ex.com/a/b/c\"); ref, _ := url.Parse(\"../d\"); _ = base.ResolveReference(ref).Path }",
    url_user_password_false_when_absent => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://guest@host/\"); _, ok := u.User.Password(); _ = ok }",
    url_user_set_via_url_user_password => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/\"); u.User = url.UserPassword(\"admin\", \"pw\"); _ = u.User.Username() }",
    url_user_set_via_url_user_only => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://host/\"); u.User = url.User(\"solo\"); _ = u.String() }",
    url_parse_opaque_path => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"mailto:user@example.com\"); _ = u.Opaque }",
    url_hostname_strips_port => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://api.test:9090/v1\"); _ = u.Hostname() }",
    url_port_extracts_numeric => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://api.test:9090/v1\"); _ = u.Port() }",
    url_escaped_path_decodes_percent => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://h/p%2Fq\"); _ = u.EscapedPath() }",
    url_request_uri_rebuild => "package main; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"/x?y=1\"); _ = u.RequestURI() }",
    url_is_abs_on_absolute_url => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://a/b\"); _ = u.IsAbs() }",
    url_is_abs_on_relative_path => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"/rel\"); _ = u.IsAbs() }",
    url_join_path_then_resolve => "package main; import \"net/url\"; func main() { base, _ := url.Parse(\"https://ex.com/a/\"); joined, _ := url.JoinPath(base.Path, \"b\"); ref, _ := url.Parse(joined); _ = base.ResolveReference(ref).Path }",
    url_query_del_nonexistent_noop => "package main; import \"net/url\"; func main() { q := url.Values{\"a\": []string{\"1\"}}; q.Del(\"z\"); _ = len(q) }",
    url_parse_request_uri_with_userinfo => "package main; import \"net/url\"; func main() { u, _ := url.ParseRequestURI(\"https://u:p@host/res\"); _ = u.User }",
    url_redacted_omits_credentials => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://secret:pass@host/path\"); _ = u.Redacted() }",
    url_join_path_preserves_root => "package main; import \"net/url\"; func main() { _, _ = url.JoinPath(\"/\", \"index.html\") }",
    url_query_add_empty_value => "package main; import \"net/url\"; func main() { q := url.Values{}; q.Add(\"flag\", \"\"); _ = q.Encode() }",
    url_path_escape_unicode_rune => "package main; import \"net/url\"; func main() { _ = url.PathEscape(\"café\") }",
    url_resolve_reference_empty_ref => "package main; import \"net/url\"; func main() { base, _ := url.Parse(\"https://ex.com/base\"); ref, _ := url.Parse(\"\"); _ = base.ResolveReference(ref).String() }",
    url_user_string_without_password => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://solo@host/\"); _ = u.User.String() }",
    url_parse_then_mutate_host => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"https://old.com/x\"); u.Host = \"new.com\"; _ = u.Host }",
    url_parse_then_mutate_scheme => "package main; import \"net/url\"; func main() { u, _ := url.Parse(\"http://h/\"); u.Scheme = \"https\"; _ = u.Scheme }",
    url_query_encode_then_parse_roundtrip => "package main; import \"net/url\"; func main() { q := url.Values{}; q.Set(\"z\", \"9\"); u, _ := url.Parse(\"https://h/?\" + q.Encode()); _ = u.Query().Get(\"z\") }",
    url_join_path_three_deep_relative => "package main; import \"net/url\"; func main() { _, _ = url.JoinPath(\"a\", \"b/c\", \"d\") }",
}
