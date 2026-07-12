crate::js_cases! {
    url_origin_omits_default_http_port => { r#"console.log(new URL("http://example.com:80/").origin);"#, ["http://example.com"] };
    url_origin_omits_default_https_port => { r#"console.log(new URL("https://example.com:443/").origin);"#, ["https://example.com"] };
    url_searchparams_live_view_reflects_search_setter => {
        r#"
const u = new URL("https://example.com/?a=1");
u.search = "?b=2";
console.log(u.searchParams.get("a") === null);
console.log(u.searchParams.get("b"));
"#,
        ["true", "2"]
    };
    url_search_setter_preserves_hash => {
        r#"
const u = new URL("https://example.com/a#x");
u.search = "?q=1";
console.log(u.href);
"#,
        ["https://example.com/a?q=1#x"]
    };
    url_hash_setter_preserves_search => {
        r#"
const u = new URL("https://example.com/a?q=1");
u.hash = "frag";
console.log(u.href);
"#,
        ["https://example.com/a?q=1#frag"]
    };
    url_pathname_empty_string_normalizes_to_slash => {
        r#"
const u = new URL("https://example.com/a");
u.pathname = "";
console.log(u.pathname);
"#,
        ["/"]
    };
    url_assigning_relative_pathname_adds_leading_slash => {
        r#"
const u = new URL("https://example.com/a");
u.pathname = "b";
console.log(u.pathname);
"#,
        ["/b"]
    };
    url_username_empty_removes_credentials_prefix_when_no_password => {
        r#"
const u = new URL("https://alice@example.com/a");
u.username = "";
console.log(u.href);
"#,
        ["https://example.com/a"]
    };
    url_password_without_username_serializes_with_empty_username_slot => {
        r#"
const u = new URL("https://example.com/a");
u.password = "secret";
console.log(u.href);
"#,
        ["https://:secret@example.com/a"]
    };
    url_host_setter_updates_hostname_and_port => {
        r#"
const u = new URL("https://example.com/a");
u.host = "api.example.com:9000";
console.log(u.hostname);
console.log(u.port);
"#,
        ["api.example.com", "9000"]
    };
    url_hostname_setter_preserves_existing_port => {
        r#"
const u = new URL("https://example.com:8080/a");
u.hostname = "api.example.com";
console.log(u.host);
"#,
        ["api.example.com:8080"]
    };
    url_port_non_numeric_assignment_is_ignored_or_cleared_to_empty_numeric_rule => {
        r#"
const u = new URL("https://example.com:8080/a");
u.port = "abc";
console.log(u.port === "8080" || u.port === "");
"#,
        ["true"]
    };
    url_searchparams_delete_updates_search_property => {
        r#"
const u = new URL("https://example.com/?a=1&b=2");
u.searchParams.delete("a");
console.log(u.search);
"#,
        ["?b=2"]
    };
    url_searchparams_append_updates_search_property => {
        r#"
const u = new URL("https://example.com/?a=1");
u.searchParams.append("b", "2");
console.log(u.search);
"#,
        ["?a=1&b=2"]
    };
    url_canparse_handles_base_with_relative_parent_segments => {
        r#"
console.log(URL.canParse("../x", "https://example.com/a/b/"));
"#,
        ["true"]
    };
    url_href_assignment_normalizes_dot_segments => {
        r#"
const u = new URL("https://example.com/a");
u.href = "https://example.com/a/../b";
console.log(u.href);
"#,
        ["https://example.com/b"]
    };
    url_stringifier_after_multiple_mutations_matches_href => {
        r#"
const u = new URL("https://example.com/a");
u.pathname = "/b";
u.search = "?q=1";
u.hash = "x";
console.log(String(u) === u.href);
"#,
        ["true"]
    };
    url_tojson_after_multiple_mutations_matches_href => {
        r#"
const u = new URL("https://example.com/a");
u.pathname = "/b";
u.search = "?q=1";
u.hash = "x";
console.log(u.toJSON());
"#,
        ["https://example.com/b?q=1#x"]
    };
    url_searchparams_set_encodes_spaces_as_plus_in_search => {
        r#"
const u = new URL("https://example.com/");
u.searchParams.set("q", "two words");
console.log(u.search);
"#,
        ["?q=two+words"]
    };
    url_pathname_preserves_encoded_space_sequence => {
        r#"
const u = new URL("https://example.com/a%20b");
console.log(u.pathname);
"#,
        ["/a%20b"]
    };
}
