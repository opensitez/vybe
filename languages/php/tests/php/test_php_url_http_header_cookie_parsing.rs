use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: URL & HTTP Utilities — parse_url, urlencode, urldecode, http_build_query, get_headers, header, setcookie
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_parse_url_components() {
    let out = run_prints(
        r#"<?php
$url = "https://user:pass@example.com:8080/path/to/script.php?param=val#section";
$parsed = parse_url($url);

echo $parsed["scheme"] . " | " . $parsed["host"] . " | " . $parsed["port"] . " | " . $parsed["path"];
"#,
    );
    assert_eq!(
        out,
        vec!["https | example.com | 8080 | /path/to/script.php"]
    );
}

#[test]
fn test_php_http_build_query_nested_arrays() {
    let out = run_prints(
        r#"<?php
$params = [
    "user" => ["name" => "Alice", "role" => "admin"],
    "filter" => "active"
];
echo http_build_query($params);
"#,
    );
    assert_eq!(
        out,
        vec!["user%5Bname%5D=Alice&user%5Brole%5D=admin&filter=active"]
    );
}

#[test]
fn test_php_urlencode_urldecode_roundtrip() {
    let out = run_prints(
        r#"<?php
$raw = "Parameter & Value with spaces / slashes";
$encoded = urlencode($raw);
$decoded = urldecode($encoded);
echo ($decoded === $raw ? "ROUNDTRIP_OK" : "ROUNDTRIP_FAIL");
"#,
    );
    assert_eq!(out, vec!["ROUNDTRIP_OK"]);
}

#[test]
fn test_php_rawurlencode_rfc3986_escaping() {
    let out = run_prints(
        r#"<?php
$str = "foo bar";
echo urlencode($str) . " vs " . rawurlencode($str);
"#,
    );
    assert_eq!(out, vec!["foo+bar vs foo%20bar"]);
}

#[test]
fn test_php_headers_list_and_sent_check() {
    compile_ok(
        r#"<?php
echo headers_sent() ? "HEADERS_SENT" : "NOT_SENT";
$headers = headers_list();
echo is_array($headers) ? "ARRAY_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_header_status_code_setting() {
    compile_ok(
        r#"<?php
if (!headers_sent()) {
    header("Content-Type: application/json; charset=UTF-8", replace: true, response_code: 200);
    header("X-Custom-Header: VybeFramework");
}
"#,
    );
}

#[test]
fn test_php_setcookie_parameters_options_array() {
    compile_ok(
        r#"<?php
if (!headers_sent()) {
    setcookie("session_id", "abc123xyz", [
        "expires" => time() + 3600,
        "path" => "/",
        "domain" => "example.com",
        "secure" => true,
        "httponly" => true,
        "samesite" => "Lax",
    ]);
}
"#,
    );
}

#[test]
fn test_php_parse_url_component_constant_filter() {
    compile_ok(
        r#"<?php
$url = "https://example.com/api/v1";
$host = parse_url($url, PHP_URL_HOST);
$path = parse_url($url, PHP_URL_PATH);
echo "$host $path";
"#,
    );
}

#[test]
fn test_php_http_response_code_get_set() {
    compile_ok(
        r#"<?php
if (!headers_sent()) {
    http_response_code(404);
    echo http_response_code();
}
"#,
    );
}

#[test]
fn test_php_header_remove_name() {
    compile_ok(
        r#"<?php
if (!headers_sent()) {
    header("X-Powered-By: Vybe");
    header_remove("X-Powered-By");
}
"#,
    );
}
