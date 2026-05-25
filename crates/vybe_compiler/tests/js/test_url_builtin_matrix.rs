crate::js_cases! {
    url_absolute_components_basic => {
        r#"
const u = new URL("https://user:pass@example.com:8080/a/b?q=1#hash");
console.log(u.protocol);
console.log(u.username);
console.log(u.password);
console.log(u.hostname);
console.log(u.port);
console.log(u.pathname);
console.log(u.search);
console.log(u.hash);
"#,
        ["https:", "user", "pass", "example.com", "8080", "/a/b", "?q=1", "#hash"]
    };

    url_origin_includes_explicit_port => {
        r#"
const u = new URL("https://example.com:8443/path");
console.log(u.origin);
"#,
        ["https://example.com:8443"]
    };

    url_host_includes_port_when_present => {
        r#"
const u = new URL("https://example.com:8443/path");
console.log(u.host);
"#,
        ["example.com:8443"]
    };

    url_host_omits_port_when_absent => {
        r#"
const u = new URL("https://example.com/path");
console.log(u.host);
"#,
        ["example.com"]
    };

    url_href_matches_stringifier => {
        r#"
const u = new URL("https://example.com/a?b=1#c");
console.log(String(u));
console.log(u.toString());
console.log(u.href);
"#,
        ["https://example.com/a?b=1#c", "https://example.com/a?b=1#c", "https://example.com/a?b=1#c"]
    };

    url_tojson_matches_href => {
        r#"
const u = new URL("https://example.com/a?b=1#c");
console.log(u.toJSON());
"#,
        ["https://example.com/a?b=1#c"]
    };

    url_canparse_absolute_true => {
        r#"
console.log(URL.canParse("https://example.com"));
"#,
        ["true"]
    };

    url_canparse_relative_with_base_true => {
        r#"
console.log(URL.canParse("child", "https://example.com/root/"));
"#,
        ["true"]
    };

    url_canparse_relative_without_base_false => {
        r#"
console.log(URL.canParse("child"));
"#,
        ["false"]
    };

    url_canparse_invalid_url_false => {
        r#"
console.log(URL.canParse("http://[:::1]"));
"#,
        ["false"]
    };

    url_relative_resolution_with_base_directory => {
        r#"
const u = new URL("child", "https://example.com/dir/");
console.log(u.href);
"#,
        ["https://example.com/dir/child"]
    };

    url_relative_resolution_with_base_file => {
        r#"
const u = new URL("child", "https://example.com/dir/file.txt");
console.log(u.href);
"#,
        ["https://example.com/dir/child"]
    };

    url_relative_resolution_normalizes_dot_segments => {
        r#"
const u = new URL("../b/./c", "https://example.com/a/d/");
console.log(u.href);
"#,
        ["https://example.com/a/b/c"]
    };

    url_relative_resolution_clamps_above_root => {
        r#"
const u = new URL("../../x", "https://example.com/a/");
console.log(u.href);
"#,
        ["https://example.com/x"]
    };

    url_protocol_setter_updates_href => {
        r#"
const u = new URL("http://example.com/a");
u.protocol = "https:";
console.log(u.href);
"#,
        ["https://example.com/a"]
    };

    url_hostname_setter_updates_host => {
        r#"
const u = new URL("https://example.com/a");
u.hostname = "api.example.com";
console.log(u.href);
"#,
        ["https://api.example.com/a"]
    };

    url_port_setter_updates_host => {
        r#"
const u = new URL("https://example.com/a");
u.port = "8443";
console.log(u.host);
"#,
        ["example.com:8443"]
    };

    url_port_empty_string_clears_port => {
        r#"
const u = new URL("https://example.com:8443/a");
u.port = "";
console.log(u.host);
"#,
        ["example.com"]
    };

    url_username_setter_adds_credentials => {
        r#"
const u = new URL("https://example.com/a");
u.username = "alice";
console.log(u.href);
"#,
        ["https://alice@example.com/a"]
    };

    url_password_setter_adds_credentials => {
        r#"
const u = new URL("https://example.com/a");
u.username = "alice";
u.password = "secret";
console.log(u.href);
"#,
        ["https://alice:secret@example.com/a"]
    };

    url_clearing_password_updates_href => {
        r#"
const u = new URL("https://alice:secret@example.com/a");
u.password = "";
console.log(u.href);
"#,
        ["https://alice@example.com/a"]
    };

    url_pathname_setter_updates_path => {
        r#"
const u = new URL("https://example.com/a");
u.pathname = "/b/c";
console.log(u.href);
"#,
        ["https://example.com/b/c"]
    };

    url_search_setter_with_question_prefix => {
        r#"
const u = new URL("https://example.com/a");
u.search = "?x=1&y=2";
console.log(u.href);
"#,
        ["https://example.com/a?x=1&y=2"]
    };

    url_search_setter_without_prefix_normalizes_prefix => {
        r#"
const u = new URL("https://example.com/a");
u.search = "x=1&y=2";
console.log(u.search);
"#,
        ["?x=1&y=2"]
    };

    url_search_empty_clears_query => {
        r#"
const u = new URL("https://example.com/a?x=1");
u.search = "";
console.log(u.href);
"#,
        ["https://example.com/a"]
    };

    url_hash_setter_with_prefix => {
        r##"
const u = new URL("https://example.com/a");
u.hash = "#part";
console.log(u.href);
    "##,
        ["https://example.com/a#part"]
    };

    url_hash_setter_without_prefix_normalizes_prefix => {
        r#"
const u = new URL("https://example.com/a");
u.hash = "part";
console.log(u.hash);
"#,
        ["#part"]
    };

    url_hash_empty_clears_fragment => {
        r##"
const u = new URL("https://example.com/a#part");
u.hash = "";
console.log(u.href);
    "##,
        ["https://example.com/a"]
    };

    url_setting_href_reparses_components => {
        r##"
const u = new URL("https://example.com/a");
u.href = "http://user:pass@other.test:8080/b?q=1#h";
console.log(u.protocol);
console.log(u.host);
console.log(u.pathname);
    "##,
        ["http:", "other.test:8080", "/b"]
    };

    url_searchparams_reads_live_query_values => {
        r#"
const u = new URL("https://example.com/a?x=1&x=2&y=3");
console.log(u.searchParams.get("y"));
console.log(u.searchParams.getAll("x").join(","));
"#,
        ["3", "1,2"]
    };

    url_searchparams_append_updates_href => {
        r#"
const u = new URL("https://example.com/a?x=1");
u.searchParams.append("x", "2");
console.log(u.href);
"#,
        ["https://example.com/a?x=1&x=2"]
    };

    url_searchparams_set_replaces_existing_values => {
        r#"
const u = new URL("https://example.com/a?x=1&x=2&y=3");
u.searchParams.set("x", "9");
console.log(u.href);
"#,
        ["https://example.com/a?x=9&y=3"]
    };

    url_searchparams_delete_removes_key => {
        r#"
const u = new URL("https://example.com/a?x=1&y=2&x=3");
u.searchParams.delete("x");
console.log(u.href);
"#,
        ["https://example.com/a?y=2"]
    };

    url_searchparams_sort_reorders_query => {
        r#"
const u = new URL("https://example.com/a?z=1&a=2&m=3");
u.searchParams.sort();
console.log(u.search);
"#,
        ["?a=2&m=3&z=1"]
    };

    url_searchparams_has_reflects_query_state => {
        r#"
const u = new URL("https://example.com/a?x=1");
console.log(u.searchParams.has("x"));
console.log(u.searchParams.has("y"));
"#,
        ["true", "false"]
    };

    url_default_port_is_omitted_in_origin => {
        r#"
const u = new URL("https://example.com:443/a");
console.log(u.origin);
console.log(u.host);
"#,
        ["https://example.com", "example.com"]
    };

    url_pathname_with_trailing_slash_roundtrips => {
        r#"
const u = new URL("https://example.com/a/b/");
console.log(u.pathname);
"#,
        ["/a/b/"]
    };

    url_searchparams_iteration_order_matches_query => {
        r#"
const u = new URL("https://example.com/a?b=1&a=2&b=3");
const out = [];
for (const [k, v] of u.searchParams) out.push(k + "=" + v);
console.log(out.join(","));
"#,
        ["b=1,a=2,b=3"]
    };

    url_protocol_property_is_lowercase => {
        r#"
const u = new URL("HTTPS://example.com/a");
console.log(u.protocol);
"#,
        ["https:"]
    };
}