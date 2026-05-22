use super::helpers::compile_ok;

// ── parse_url full decomposition ─────────────────────────────────
#[test]
fn parse_url_full_url() { compile_ok(r#"<?php
$parts = parse_url('https://user:pass@example.com:8080/path?q=1#frag');
echo $parts['scheme'];
echo $parts['host'];
echo $parts['port'];
"#); }

// ── parse_url PHP_URL_HOST component ────────────────────────────
#[test]
fn parse_url_component_host() { compile_ok(r#"<?php
$host = parse_url('https://example.com/page', PHP_URL_HOST);
echo $host;
"#); }

// ── parse_url PHP_URL_PATH component ────────────────────────────
#[test]
fn parse_url_component_path() { compile_ok(r#"<?php
$path = parse_url('https://example.com/foo/bar?x=1', PHP_URL_PATH);
echo $path;
"#); }

// ── parse_url PHP_URL_QUERY component ───────────────────────────
#[test]
fn parse_url_component_query() { compile_ok(r#"<?php
$query = parse_url('https://example.com/search?q=hello&page=2', PHP_URL_QUERY);
echo $query;
"#); }

// ── parse_url PHP_URL_PORT component ────────────────────────────
#[test]
fn parse_url_component_port() { compile_ok(r#"<?php
$port = parse_url('http://example.com:3000/app', PHP_URL_PORT);
echo $port;
"#); }

// ── http_build_query from associative array ──────────────────────
#[test]
fn http_build_query_assoc_array() { compile_ok(r#"<?php
$params = ['name' => 'Alice', 'age' => 30, 'city' => 'Paris'];
$qs = http_build_query($params);
echo $qs;
"#); }

// ── http_build_query with custom separator ───────────────────────
#[test]
fn http_build_query_custom_separator() { compile_ok(r#"<?php
$params = ['a' => 1, 'b' => 2, 'c' => 3];
$qs = http_build_query($params, '', '&amp;');
echo $qs;
"#); }

// ── http_build_query with numeric array ─────────────────────────
#[test]
fn http_build_query_numeric_array() { compile_ok(r#"<?php
$params = ['foo', 'bar', 'baz'];
$qs = http_build_query($params);
echo $qs;
"#); }

// ── parse_str query string into variables ────────────────────────
#[test]
fn parse_str_into_array() { compile_ok(r#"<?php
parse_str('name=Bob&score=100&active=1', $out);
echo $out['name'];
echo $out['score'];
echo $out['active'];
"#); }

// ── urlencode vs rawurlencode ────────────────────────────────────
#[test]
fn urlencode_vs_rawurlencode() { compile_ok(r#"<?php
$str = 'hello world & more';
echo urlencode($str);
echo rawurlencode($str);
"#); }

// ── urldecode vs rawurldecode ────────────────────────────────────
#[test]
fn urldecode_vs_rawurldecode() { compile_ok(r#"<?php
$encoded = 'hello+world%20%26%20more';
echo urldecode($encoded);
echo rawurldecode($encoded);
"#); }

// ── base64 encode/decode roundtrip ──────────────────────────────
#[test]
fn base64_encode_decode_roundtrip() { compile_ok(r#"<?php
$original = 'Hello, World! This is a test string.';
$encoded = base64_encode($original);
$decoded = base64_decode($encoded);
echo $decoded === $original ? 'match' : 'mismatch';
echo $encoded;
"#); }

// ── http_response_code ───────────────────────────────────────────
#[test]
fn http_response_code_compile_ok() { compile_ok(r#"<?php
$code = http_response_code();
echo is_int($code) || $code === false ? 'ok' : 'fail';
"#); }

// ── headers_sent ─────────────────────────────────────────────────
#[test]
fn headers_sent_compile_ok() { compile_ok(r#"<?php
$sent = headers_sent();
echo is_bool($sent) ? 'bool' : 'other';
"#); }

// ── header send ──────────────────────────────────────────────────
#[test]
fn header_send_compile_ok() { compile_ok(r#"<?php
header('Content-Type: application/json');
header('X-Custom-Header: value');
echo 'done';
"#); }

// ── setcookie ────────────────────────────────────────────────────
#[test]
fn setcookie_compile_ok() { compile_ok(r#"<?php
setcookie('session_id', 'abc123', time() + 3600, '/', '', true, true);
echo 'cookie set';
"#); }

// ── ip2long ──────────────────────────────────────────────────────
#[test]
fn ip2long_ipv4_address() { compile_ok(r#"<?php
$n = ip2long('192.168.1.1');
echo $n > 0 ? 'positive' : 'fail';
$loopback = ip2long('127.0.0.1');
echo $loopback;
"#); }

// ── long2ip ──────────────────────────────────────────────────────
#[test]
fn long2ip_integer_to_ip() { compile_ok(r#"<?php
$ip = long2ip(2130706433);
echo $ip;
$roundtrip = long2ip(ip2long('10.0.0.1'));
echo $roundtrip;
"#); }

// ── inet_pton ────────────────────────────────────────────────────
#[test]
fn inet_pton_compile_ok() { compile_ok(r#"<?php
$packed = inet_pton('127.0.0.1');
echo $packed !== false ? 'ok' : 'fail';
$packed6 = inet_pton('::1');
echo $packed6 !== false ? 'ok' : 'fail';
"#); }

// ── gethostbyname ────────────────────────────────────────────────
#[test]
fn gethostbyname_compile_ok() { compile_ok(r#"<?php
$ip = gethostbyname('localhost');
echo is_string($ip) ? 'string' : 'other';
"#); }
