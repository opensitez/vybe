//! Runtime behavior for URL/HTTP helpers (host-wired subset of `test_url_http.rs`).

crate::php_cases! {
    urlencode_spaces_and_ampersand => {
        r#"<?php
echo urlencode('hello world & more');
"#,
        ["hello+world+%26+more"]
    };

    rawurlencode_spaces_and_ampersand => {
        r#"<?php
echo rawurlencode('hello world & more');
"#,
        ["hello%20world%20%26%20more"]
    };

    urldecode_plus_and_percent => {
        r#"<?php
echo urldecode('hello+world%20%26%20more');
"#,
        ["hello world & more"]
    };

    rawurldecode_percent_encoding => {
        r#"<?php
echo rawurldecode('hello+world%20%26%20more');
"#,
        ["hello+world & more"]
    };

    base64_encode_decode_roundtrip => {
        r#"<?php
$original = 'Hello, World! This is a test string.';
$encoded = base64_encode($original);
$decoded = base64_decode($encoded);
echo ($decoded === $original ? 'match' : 'mismatch') . ':' . $encoded;
"#,
        ["match:SGVsbG8sIFdvcmxkISBUaGlzIGlzIGEgdGVzdCBzdHJpbmcu"]
    };

    headers_sent_returns_bool => {
        r#"<?php
$sent = headers_sent();
echo is_bool($sent) ? 'bool' : 'other';
"#,
        ["bool"]
    };

    header_send_then_echo_done => {
        r#"<?php
header('Content-Type: application/json');
header('X-Custom-Header: value');
echo 'done';
"#,
        ["done"]
    };

    setcookie_then_echo_confirmation => {
        r#"<?php
setcookie('session_id', 'abc123', time() + 3600, '/', '', true, true);
echo 'cookie set';
"#,
        ["cookie set"]
    };

    http_response_code_returns_int_or_false => {
        r#"<?php
$code = http_response_code();
echo is_int($code) || $code === false ? 'ok' : 'fail';
"#,
        ["ok"]
    };
}
