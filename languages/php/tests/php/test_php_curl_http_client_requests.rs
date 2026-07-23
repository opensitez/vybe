use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: cURL HTTP Client Interface — curl_init, curl_setopt, curl_exec, curl_getinfo, curl_errno, curl_error
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_curl_init_setopt_get_info_constants() {
    let out = run_prints(
        r#"<?php
$ch = curl_init("https://example.com/api/test");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_TIMEOUT, 10);
curl_setopt($ch, CURLOPT_HTTPHEADER, ["Accept: application/json"]);

$info = curl_getinfo($ch);
curl_close($ch);

echo "URL={$info['url']} Opts_Set=1";
"#,
    );
    assert_eq!(out, vec!["URL=https://example.com/api/test Opts_Set=1"]);
}

#[test]
fn test_php_curl_setopt_array_batch_configuration() {
    let out = run_prints(
        r#"<?php
$ch = curl_init();
$options = [
    CURLOPT_URL => "https://api.github.com/users",
    CURLOPT_USERAGENT => "Vybe-Client/1.0",
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_HEADER => false,
];

curl_setopt_array($ch, $options);
$url = curl_getinfo($ch, CURLINFO_EFFECTIVE_URL);
curl_close($ch);

echo "Effective URL: $url";
"#,
    );
    assert_eq!(out, vec!["Effective URL: https://api.github.com/users"]);
}

#[test]
fn test_php_curl_error_and_errno_handling() {
    let out = run_prints(
        r#"<?php
$ch = curl_init("http://invalid.domain.nonexistent.vybe");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_TIMEOUT, 1);
@curl_exec($ch);

$errno = curl_errno($ch);
curl_close($ch);

echo is_int($errno) ? "ERRNO_IS_INT" : "FAIL";
"#,
    );
    assert_eq!(out, vec!["ERRNO_IS_INT"]);
}

#[test]
fn test_php_curl_version_info_structure() {
    compile_ok(
        r#"<?php
$v = curl_version();
echo "cURL Version: " . $v["version"] . " SSL: " . $v["ssl_version"];
"#,
    );
}

#[test]
fn test_php_curl_post_fields_array_multipart() {
    compile_ok(
        r#"<?php
$ch = curl_init("https://httpbin.org/post");
curl_setopt($ch, CURLOPT_POST, true);
curl_setopt($ch, CURLOPT_POSTFIELDS, [
    "username" => "alice",
    "file" => "data_content"
]);
curl_close($ch);
"#,
    );
}

#[test]
fn test_php_curl_multi_init_add_remove_handle() {
    compile_ok(
        r#"<?php
$mh = curl_multi_init();
$ch1 = curl_init("https://example.com/1");
$ch2 = curl_init("https://example.com/2");

curl_multi_add_handle($mh, $ch1);
curl_multi_add_handle($mh, $ch2);

curl_multi_remove_handle($mh, $ch1);
curl_multi_remove_handle($mh, $ch2);
curl_multi_close($mh);
curl_close($ch1);
curl_close($ch2);
echo "MULTI_CURL_OK";
"#,
    );
}

#[test]
fn test_php_curl_reset_handle_options() {
    compile_ok(
        r#"<?php
$ch = curl_init("https://example.com");
curl_setopt($ch, CURLOPT_TIMEOUT, 5);
curl_reset($ch);
$url = curl_getinfo($ch, CURLINFO_EFFECTIVE_URL);
curl_close($ch);
echo "Reset OK";
"#,
    );
}

#[test]
fn test_php_curl_copy_handle_duplication() {
    compile_ok(
        r#"<?php
$ch = curl_init("https://example.com");
curl_setopt($ch, CURLOPT_USERAGENT, "TestAgent");
$copy = curl_copy_handle($ch);
curl_close($ch);
curl_close($copy);
echo "Copy OK";
"#,
    );
}

#[test]
fn test_php_curl_strerror_error_string() {
    compile_ok(
        r#"<?php
$errStr = curl_strerror(CURLE_COULDNT_RESOLVE_HOST);
echo is_string($errStr) && strlen($errStr) > 0 ? "STRERROR_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_curl_escape_unescape_urls() {
    compile_ok(
        r#"<?php
$ch = curl_init();
$escaped = curl_escape($ch, "hello world & php");
$unescaped = curl_unescape($ch, $escaped);
curl_close($ch);
echo $unescaped === "hello world & php" ? "ESCAPE_ROUNDTRIP_OK" : "FAIL";
"#,
    );
}
