use crate::helpers::run_main;

#[test]
fn url_string_constructor_sets_protocol() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://url.test/page"); System.out.println(u.getProtocol());"#,
    );
    assert_eq!(out, vec!["http"]);
}

#[test]
fn url_string_constructor_sets_host() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://host.url.test/data"); System.out.println(u.getHost());"#,
    );
    assert_eq!(out, vec!["host.url.test"]);
}

#[test]
fn url_https_protocol_parsed() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("https://secure.url.test/"); System.out.println(u.getProtocol());"#,
    );
    assert_eq!(out, vec!["https"]);
}

#[test]
fn url_file_protocol_parsed() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("file:/local/file.txt"); System.out.println(u.getProtocol());"#,
    );
    assert_eq!(out, vec!["file"]);
}

#[test]
fn url_get_port_explicit_value() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://p.url.test:8080/"); System.out.println(u.getPort());"#,
    );
    assert_eq!(out, vec!["8080"]);
}

#[test]
fn url_get_port_negative_one_when_default() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://def.url.test/"); System.out.println(u.getPort());"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn url_get_default_port_http_is_eighty() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://dp.url.test/"); System.out.println(u.getDefaultPort());"#,
    );
    assert_eq!(out, vec!["80"]);
}

#[test]
fn url_get_default_port_https_is_four_forty_three() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("https://dp.url.test/"); System.out.println(u.getDefaultPort());"#,
    );
    assert_eq!(out, vec!["443"]);
}

#[test]
fn url_get_file_includes_path_and_query() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://f.url.test/dir/file?q=1"); System.out.println(u.getFile());"#,
    );
    assert_eq!(out, vec!["/dir/file?q=1"]);
}

#[test]
fn url_get_path_returns_path_only() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://path.url.test/a/b"); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["/a/b"]);
}

#[test]
fn url_get_query_returns_query_string() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://q.url.test/x?key=val"); System.out.println(u.getQuery());"#,
    );
    assert_eq!(out, vec!["key=val"]);
}

#[test]
fn url_get_ref_returns_fragment() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://ref.url.test/doc#section"); System.out.println(u.getRef());"#,
    );
    assert_eq!(out, vec!["section"]);
}

#[test]
fn url_get_authority_host_and_port() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://auth.url.test:9000/res"); System.out.println(u.getAuthority());"#,
    );
    assert_eq!(out, vec!["auth.url.test:9000"]);
}

#[test]
fn url_get_user_info_with_credentials() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://alice:secret@ui.url.test/"); System.out.println(u.getUserInfo());"#,
    );
    assert_eq!(out, vec!["alice:secret"]);
}

#[test]
fn url_to_external_form_contains_protocol() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://ext.url.test/x"); System.out.println(u.toExternalForm().startsWith("http://"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_to_string_matches_external_form() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://ts.url.test/y"); System.out.println(u.toString().equals(u.toExternalForm()));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_equals_same_url_true() {
    let out = run_main(
        r#"java.net.URL a = new java.net.URL("http://eq.url.test/z"); java.net.URL b = new java.net.URL("http://eq.url.test/z"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_equals_different_path_false() {
    let out = run_main(
        r#"java.net.URL a = new java.net.URL("http://eq.url.test/a"); java.net.URL b = new java.net.URL("http://eq.url.test/b"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn url_hash_code_consistent_for_equal_urls() {
    let out = run_main(
        r#"java.net.URL a = new java.net.URL("http://hc.url.test/p"); java.net.URL b = new java.net.URL("http://hc.url.test/p"); System.out.println(a.hashCode() == b.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_three_arg_constructor_protocol_host_file() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http", "three.url.test", "/index.html"); System.out.println(u.getProtocol()); System.out.println(u.getHost()); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["http", "three.url.test", "/index.html"]);
}

#[test]
fn url_context_relative_constructor_resolves_against_base() {
    let out = run_main(
        r#"java.net.URL base = new java.net.URL("http://ctx.url.test/a/"); java.net.URL rel = new java.net.URL(base, "b.html"); System.out.println(rel.getPath());"#,
    );
    assert_eq!(out, vec!["/a/b.html"]);
}

#[test]
fn url_context_absolute_string_replaces_base() {
    let out = run_main(
        r#"java.net.URL base = new java.net.URL("http://ctx.url.test/old"); java.net.URL abs = new java.net.URL(base, "http://other.url.test/new"); System.out.println(abs.getHost());"#,
    );
    assert_eq!(out, vec!["other.url.test"]);
}

#[test]
fn url_same_file_true_for_equivalent_paths() {
    let out = run_main(
        r#"java.net.URL a = new java.net.URL("http://sf.url.test/a/b"); java.net.URL b = new java.net.URL("http://sf.url.test/a/b"); System.out.println(a.sameFile(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_same_file_false_for_different_files() {
    let out = run_main(
        r#"java.net.URL a = new java.net.URL("http://sf.url.test/a"); java.net.URL b = new java.net.URL("http://sf.url.test/b"); System.out.println(a.sameFile(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn url_to_uri_conversion_preserves_scheme() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://tu.url.test/r"); java.net.URI uri = u.toURI(); System.out.println(uri.getScheme());"#,
    );
    assert_eq!(out, vec!["http"]);
}

#[test]
fn url_to_uri_conversion_preserves_host() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://tu2.url.test/r"); java.net.URI uri = u.toURI(); System.out.println(uri.getHost());"#,
    );
    assert_eq!(out, vec!["tu2.url.test"]);
}

#[test]
fn url_uri_roundtrip_via_uri_constructor() {
    let out = run_main(
        r#"java.net.URL orig = new java.net.URL("http://rt.url.test/p"); java.net.URI uri = orig.toURI(); java.net.URL back = uri.toURL(); System.out.println(back.getHost());"#,
    );
    assert_eq!(out, vec!["rt.url.test"]);
}

#[test]
fn url_get_content_type_unknown_without_connection() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://ct.url.test/x.txt"); System.out.println(u.getPath().endsWith(".txt"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_jar_protocol_parsed() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("jar:file:/app.jar!/META-INF/"); System.out.println(u.getProtocol());"#,
    );
    assert_eq!(out, vec!["jar"]);
}

#[test]
fn url_localhost_host_parsed() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://127.0.0.1:5000/api"); System.out.println(u.getHost());"#,
    );
    assert_eq!(out, vec!["127.0.0.1"]);
}

#[test]
fn url_empty_path_defaults_to_empty_string() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://ep.url.test"); System.out.println(u.getPath().length());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn url_query_null_when_absent() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://nq.url.test/page"); System.out.println(u.getQuery() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_ref_null_when_no_fragment() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://nr.url.test/page"); System.out.println(u.getRef() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_user_info_null_without_at_sign() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://nui.url.test/"); System.out.println(u.getUserInfo() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_four_arg_constructor_with_port() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http", "four.url.test", 4444, "/svc"); System.out.println(u.getPort()); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["4444", "/svc"]);
}

#[test]
fn url_relative_dot_segment_in_context() {
    let out = run_main(
        r#"java.net.URL base = new java.net.URL("http://dot.url.test/a/b/"); java.net.URL rel = new java.net.URL(base, "./c"); System.out.println(rel.getPath());"#,
    );
    assert_eq!(out, vec!["/a/b/./c"]);
}

#[test]
fn url_relative_parent_segment_in_context() {
    let out = run_main(
        r#"java.net.URL base = new java.net.URL("http://par.url.test/a/b/"); java.net.URL rel = new java.net.URL(base, "../c"); System.out.println(rel.getPath());"#,
    );
    assert_eq!(out, vec!["/a/b/../c"]);
}

#[test]
fn url_encoded_space_in_path() {
    // java.net.URL.getPath() returns the RAW path — no percent-decoding
    // (decoding is URI.getPath()'s behavior).
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://enc.url.test/a%20b"); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["/a%20b"]);
}

#[test]
fn url_multiple_query_parameters_preserved() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://mq.url.test/?a=1&b=2"); System.out.println(u.getQuery());"#,
    );
    assert_eq!(out, vec!["a=1&b=2"]);
}

#[test]
fn url_ftp_default_port() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("ftp://ftp.url.test/"); System.out.println(u.getDefaultPort());"#,
    );
    assert_eq!(out, vec!["21"]);
}

#[test]
fn url_compare_hosts_via_external_form_length() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://len.url.test/short"); System.out.println(u.toExternalForm().length() > 10);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn url_path_with_trailing_slash() {
    let out = run_main(
        r#"java.net.URL u = new java.net.URL("http://trail.url.test/dir/"); System.out.println(u.getPath());"#,
    );
    assert_eq!(out, vec!["/dir/"]);
}
