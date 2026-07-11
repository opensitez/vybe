use crate::helpers::run_main;

#[test]
fn uri_create_http_scheme_is_http() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://example.com"); System.out.println(u.getScheme());"#,
    );
    assert_eq!(out, vec!["http"]);
}

#[test]
fn uri_create_https_scheme_is_https() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("https://secure.example.org"); System.out.println(u.getScheme());"#,
    );
    assert_eq!(out, vec!["https"]);
}

#[test]
fn uri_create_file_scheme_is_file() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("file:///tmp/data.txt"); System.out.println(u.getScheme());"#,
    );
    assert_eq!(out, vec!["file"]);
}

#[test]
fn uri_create_ftp_scheme_is_ftp() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("ftp://files.example.net/pub"); System.out.println(u.getScheme());"#,
    );
    assert_eq!(out, vec!["ftp"]);
}

#[test]
fn uri_get_host_returns_domain_name() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://api.example.com/path"); System.out.println(u.getHost());"#,
    );
    assert_eq!(out, vec!["api.example.com"]);
}

#[test]
fn uri_get_host_returns_localhost() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://localhost:8080/"); System.out.println(u.getHost());"#,
    );
    assert_eq!(out, vec!["localhost"]);
}

#[test]
fn uri_get_port_returns_explicit_port_number() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://host.example:9090/"); System.out.println(u.getPort());"#,
    );
    assert_eq!(out, vec!["9090"]);
}

#[test]
fn uri_get_port_returns_negative_one_when_absent() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://host.example/"); System.out.println(u.getPort());"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn uri_get_path_returns_resource_path() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/a/b/c"); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["/a/b/c"]);
}

#[test]
fn uri_get_path_root_only_is_slash() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/"); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["/"]);
}

#[test]
fn uri_get_path_empty_when_no_slash_in_hierarchical() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com"); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec![""]);
}

#[test]
fn uri_get_query_returns_parameter_string() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/search?q=vybe&lang=java"); System.out.println(u.getQuery());"#,
    );
    assert_eq!(out, vec!["q=vybe&lang=java"]);
}

#[test]
fn uri_get_query_null_when_absent() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/page"); System.out.println(u.getQuery() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_get_fragment_returns_anchor() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://docs.com/guide#intro"); System.out.println(u.getFragment());"#,
    );
    assert_eq!(out, vec!["intro"]);
}

#[test]
fn uri_get_fragment_null_when_absent() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://docs.com/guide"); System.out.println(u.getFragment() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_get_authority_host_only() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://server.test/res"); System.out.println(u.getAuthority());"#,
    );
    assert_eq!(out, vec!["server.test"]);
}

#[test]
fn uri_get_authority_includes_explicit_port() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://server.test:3000/res"); System.out.println(u.getAuthority());"#,
    );
    assert_eq!(out, vec!["server.test:3000"]);
}

#[test]
fn uri_get_user_info_returns_credentials() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://user:pass@auth.test/"); System.out.println(u.getUserInfo());"#,
    );
    assert_eq!(out, vec!["user:pass"]);
}

#[test]
fn uri_get_user_info_null_without_at_sign() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://plain.test/"); System.out.println(u.getUserInfo() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_get_raw_path_preserves_percent_encoding() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/a%20b"); System.out.println(u.getRawPath());"#,
    );
    assert_eq!(out, vec!["/a%20b"]);
}

#[test]
fn uri_get_raw_query_preserves_plus_sign() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/?a+b"); System.out.println(u.getRawQuery());"#,
    );
    assert_eq!(out, vec!["a+b"]);
}

#[test]
fn uri_get_raw_fragment_preserves_encoding() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/#sec%201"); System.out.println(u.getRawFragment());"#,
    );
    assert_eq!(out, vec!["sec%201"]);
}

#[test]
fn uri_is_absolute_true_for_http_uri() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://abs.test/"); System.out.println(u.isAbsolute());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_is_absolute_false_for_relative_path_only() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("/relative/path"); System.out.println(u.isAbsolute());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uri_is_opaque_true_for_mailto() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("mailto:dev@example.com"); System.out.println(u.isOpaque());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_is_opaque_true_for_urn() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("urn:isbn:0451450523"); System.out.println(u.isOpaque());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_is_opaque_false_for_hierarchical_http() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://hier.test/a"); System.out.println(u.isOpaque());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uri_normalize_collapses_dot_dot_segment() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/a/b/../c"); System.out.println(u.normalize().getPath());"#,
    );
    assert_eq!(out, vec!["/a/c"]);
}

#[test]
fn uri_normalize_removes_single_dot_segment() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://x.com/a/./b"); System.out.println(u.normalize().getPath());"#,
    );
    assert_eq!(out, vec!["/a/b"]);
}

#[test]
fn uri_resolve_absolute_base_with_relative_path() {
    let out = run_main(
        r#"java.net.URI base = java.net.URI.create("http://x.com/a/b"); java.net.URI rel = java.net.URI.create("c"); System.out.println(base.resolve(rel).getPath());"#,
    );
    assert_eq!(out, vec!["/a/c"]);
}

#[test]
fn uri_resolve_relative_ref_against_directory_base() {
    let out = run_main(
        r#"java.net.URI base = java.net.URI.create("http://x.com/dir/"); java.net.URI rel = java.net.URI.create("item"); System.out.println(base.resolve(rel).getPath());"#,
    );
    assert_eq!(out, vec!["/dir/item"]);
}

#[test]
fn uri_resolve_absolute_ref_replaces_entire_uri() {
    let out = run_main(
        r#"java.net.URI base = java.net.URI.create("http://old.com/a"); java.net.URI abs = java.net.URI.create("http://new.com/b"); System.out.println(base.resolve(abs).getHost());"#,
    );
    assert_eq!(out, vec!["new.com"]);
}

#[test]
fn uri_relativize_strips_common_prefix() {
    let out = run_main(
        r#"java.net.URI a = java.net.URI.create("http://x.com/a/b"); java.net.URI b = java.net.URI.create("http://x.com/a/c"); System.out.println(a.relativize(b).getPath());"#,
    );
    assert_eq!(out, vec!["../c"]);
}

#[test]
fn uri_equals_true_for_identical_uris() {
    let out = run_main(
        r#"java.net.URI a = java.net.URI.create("http://eq.test/x"); java.net.URI b = java.net.URI.create("http://eq.test/x"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_equals_false_for_different_hosts() {
    let out = run_main(
        r#"java.net.URI a = java.net.URI.create("http://one.test/x"); java.net.URI b = java.net.URI.create("http://two.test/x"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uri_hash_code_equal_for_equal_uris() {
    let out = run_main(
        r#"java.net.URI a = java.net.URI.create("http://hc.test/p"); java.net.URI b = java.net.URI.create("http://hc.test/p"); System.out.println(a.hashCode() == b.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_to_string_contains_scheme_and_host() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://str.test/path"); String s = u.toString(); System.out.println(s.startsWith("http://str.test"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_to_ascii_string_renders_uri() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://ascii.test/x"); System.out.println(u.toASCIIString().length() > 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_compare_to_orders_by_scheme_lexicographically() {
    let out = run_main(
        r#"java.net.URI a = java.net.URI.create("http://x.com"); java.net.URI b = java.net.URI.create("https://x.com"); System.out.println(a.compareTo(b) < 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_compare_to_zero_for_equal_uris() {
    let out = run_main(
        r#"java.net.URI a = java.net.URI.create("http://cmp.test/a"); java.net.URI b = java.net.URI.create("http://cmp.test/a"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn uri_get_scheme_specific_part_opaque_mailto() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("mailto:a@b.c"); System.out.println(u.getSchemeSpecificPart());"#,
    );
    assert_eq!(out, vec!["a@b.c"]);
}

#[test]
fn uri_get_scheme_specific_part_hierarchical_includes_authority() {
    let out = run_main(
        r#"java.net.URI u = java.net.URI.create("http://ssp.test/p"); String ssp = u.getSchemeSpecificPart(); System.out.println(ssp.contains("ssp.test"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uri_seven_arg_constructor_sets_scheme_and_path() {
    let out = run_main(
        r#"java.net.URI u = new java.net.URI("http", "user", "host.test", 80, "/api", "q=1", "frag"); System.out.println(u.getScheme()); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["http", "/api"]);
}

#[test]
fn uri_three_arg_constructor_scheme_host_path() {
    let out = run_main(
        r#"java.net.URI u = new java.net.URI("http", "three.test", "/v"); System.out.println(u.getHost()); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["three.test", "/v"]);
}

#[test]
fn uri_single_string_constructor_parses_full_uri() {
    let out = run_main(
        r#"java.net.URI u = new java.net.URI("https://single.test/resource"); System.out.println(u.getScheme()); System.out.println(u.getHost());"#,
    );
    assert_eq!(out, vec!["https", "single.test"]);
}
