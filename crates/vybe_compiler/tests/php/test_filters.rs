use super::helpers::compile_ok;

// ── FILTER_VALIDATE_* ────────────────────────────────────────

#[test]
fn filter_validate_email() {
    compile_ok(
        r#"<?php
$emails = ['user@example.com', 'invalid-email', 'a@b.c', 'missing@', '@nodomain.com'];
foreach ($emails as $e) {
    echo filter_var($e, FILTER_VALIDATE_EMAIL) !== false ? 'valid' : 'invalid';
    echo ' ';
}
"#,
    );
}

#[test]
fn filter_validate_url() {
    compile_ok(
        r#"<?php
$urls = [
    'https://example.com',
    'http://foo.bar/path?query=1',
    'ftp://files.server.com',
    'not a url',
    '//relative.com',
];
foreach ($urls as $u) {
    echo filter_var($u, FILTER_VALIDATE_URL) !== false ? 'valid' : 'invalid';
    echo ' ';
}
"#,
    );
}

#[test]
fn filter_validate_ip() {
    compile_ok(
        r#"<?php
$ips = ['192.168.1.1', '256.0.0.1', '::1', '2001:db8::1', 'not-an-ip'];
foreach ($ips as $ip) {
    echo filter_var($ip, FILTER_VALIDATE_IP) !== false ? 'valid' : 'invalid';
    echo ' ';
}
"#,
    );
}

#[test]
fn filter_validate_ip_v4_only() {
    compile_ok(
        r#"<?php
echo filter_var('192.168.1.1', FILTER_VALIDATE_IP, FILTER_FLAG_IPV4) !== false ? 'v4' : 'fail';
echo filter_var('::1',          FILTER_VALIDATE_IP, FILTER_FLAG_IPV4) !== false ? 'v4' : 'fail';
"#,
    );
}

#[test]
fn filter_validate_ip_v6_only() {
    compile_ok(
        r#"<?php
echo filter_var('::1',         FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) !== false ? 'v6' : 'fail';
echo filter_var('192.168.1.1', FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) !== false ? 'v6' : 'fail';
"#,
    );
}

#[test]
fn filter_validate_int_basic() {
    compile_ok(
        r#"<?php
var_dump(filter_var('42',   FILTER_VALIDATE_INT));
var_dump(filter_var('-10',  FILTER_VALIDATE_INT));
var_dump(filter_var('3.14', FILTER_VALIDATE_INT));
var_dump(filter_var('abc',  FILTER_VALIDATE_INT));
"#,
    );
}

#[test]
fn filter_validate_int_range() {
    compile_ok(
        r#"<?php
$options = ['options' => ['min_range' => 1, 'max_range' => 100]];
var_dump(filter_var('50',  FILTER_VALIDATE_INT, $options));
var_dump(filter_var('0',   FILTER_VALIDATE_INT, $options));
var_dump(filter_var('101', FILTER_VALIDATE_INT, $options));
"#,
    );
}

#[test]
fn filter_validate_float() {
    compile_ok(
        r#"<?php
var_dump(filter_var('3.14',   FILTER_VALIDATE_FLOAT));
var_dump(filter_var('1e5',    FILTER_VALIDATE_FLOAT));
var_dump(filter_var('-0.001', FILTER_VALIDATE_FLOAT));
var_dump(filter_var('abc',    FILTER_VALIDATE_FLOAT));
"#,
    );
}

#[test]
fn filter_validate_boolean() {
    compile_ok(
        r#"<?php
$trues  = ['1', 'true', 'on', 'yes'];
$falses = ['0', 'false', 'off', 'no', ''];
foreach ($trues as $v) {
    echo filter_var($v, FILTER_VALIDATE_BOOLEAN) ? 't' : 'f';
}
echo ':';
foreach ($falses as $v) {
    echo filter_var($v, FILTER_VALIDATE_BOOLEAN) ? 't' : 'f';
}
"#,
    );
}

#[test]
fn filter_validate_domain() {
    compile_ok(
        r#"<?php
echo filter_var('example.com',     FILTER_VALIDATE_DOMAIN) !== false ? 'valid' : 'invalid';
echo filter_var('sub.domain.co.uk',FILTER_VALIDATE_DOMAIN) !== false ? 'valid' : 'invalid';
echo filter_var('invalid_domain',  FILTER_VALIDATE_DOMAIN) !== false ? 'valid' : 'invalid';
"#,
    );
}

// ── FILTER_SANITIZE_* ────────────────────────────────────────

#[test]
fn filter_sanitize_string() {
    compile_ok(
        r#"<?php
$raw = '<script>alert("xss")</script>Hello';
$clean = filter_var($raw, FILTER_SANITIZE_SPECIAL_CHARS);
echo strlen($clean) > 0 ? 'sanitized' : 'empty';
"#,
    );
}

#[test]
fn filter_sanitize_email() {
    compile_ok(
        r#"<?php
$email = "user name@exa mple.com";
$clean = filter_var($email, FILTER_SANITIZE_EMAIL);
echo $clean;
"#,
    );
}

#[test]
fn filter_sanitize_url() {
    compile_ok(
        r#"<?php
$url = "https://example.com/path with spaces?q=hello world";
$clean = filter_var($url, FILTER_SANITIZE_URL);
echo str_contains($clean, 'example.com') ? 'ok' : 'fail';
"#,
    );
}

#[test]
fn filter_sanitize_number_int() {
    compile_ok(
        r#"<?php
echo filter_var('  42abc-5  ', FILTER_SANITIZE_NUMBER_INT);
"#,
    );
}

#[test]
fn filter_sanitize_number_float() {
    compile_ok(
        r#"<?php
echo filter_var('1,234.56abc', FILTER_SANITIZE_NUMBER_FLOAT, FILTER_FLAG_ALLOW_FRACTION);
"#,
    );
}

// ── filter_var with options ───────────────────────────────────

#[test]
fn filter_var_default_option() {
    compile_ok(
        r#"<?php
$result = filter_var('', FILTER_VALIDATE_INT, ['options' => ['default' => -1]]);
echo $result;
"#,
    );
}

#[test]
fn filter_var_callback() {
    compile_ok(
        r#"<?php
$result = filter_var('hello world', FILTER_CALLBACK, ['options' => 'strtoupper']);
echo $result;
"#,
    );
}

// ── filter_var_array ─────────────────────────────────────────

#[test]
fn filter_var_array_basic() {
    compile_ok(
        r#"<?php
$data = [
    'age'   => '25',
    'email' => 'user@example.com',
    'score' => '3.14',
];
$filters = [
    'age'   => FILTER_VALIDATE_INT,
    'email' => FILTER_VALIDATE_EMAIL,
    'score' => FILTER_VALIDATE_FLOAT,
];
$result = filter_var_array($data, $filters);
var_dump($result['age']);
echo $result['email'] !== false ? 'valid email' : 'invalid email';
"#,
    );
}

#[test]
fn filter_var_array_with_options() {
    compile_ok(
        r#"<?php
$data = ['quantity' => '5', 'price' => '9.99'];
$filters = [
    'quantity' => ['filter'  => FILTER_VALIDATE_INT,
                   'options' => ['min_range' => 1, 'max_range' => 100]],
    'price'    => FILTER_VALIDATE_FLOAT,
];
$result = filter_var_array($data, $filters);
echo $result['quantity'] . ':' . $result['price'];
"#,
    );
}

// ── FILTER_DEFAULT / FILTER_UNSAFE_RAW ───────────────────────

#[test]
fn filter_default_passthrough() {
    compile_ok(
        r#"<?php
$raw = "hello <world>";
echo filter_var($raw, FILTER_DEFAULT) === $raw ? 'passthrough' : 'changed';
"#,
    );
}

// ── filter_has_var ────────────────────────────────────────────

#[test]
fn filter_has_var_env() {
    compile_ok(
        r#"<?php
// INPUT_ENV checks environment
$has = filter_has_var(INPUT_ENV, 'PATH');
echo is_bool($has) ? 'bool result' : 'non-bool';
"#,
    );
}

// ── Practical validation patterns ────────────────────────────

#[test]
fn validate_form_input() {
    compile_ok(
        r#"<?php
function validateUser(array $input): array {
    $errors = [];
    if (filter_var($input['email'] ?? '', FILTER_VALIDATE_EMAIL) === false) {
        $errors[] = 'Invalid email';
    }
    if (filter_var($input['age'] ?? '', FILTER_VALIDATE_INT,
        ['options' => ['min_range' => 0, 'max_range' => 150]]) === false) {
        $errors[] = 'Invalid age';
    }
    return $errors;
}
$valid   = validateUser(['email' => 'bob@example.com', 'age' => '25']);
$invalid = validateUser(['email' => 'notanemail', 'age' => '200']);
echo count($valid) . ':' . count($invalid);
"#,
    );
}

#[test]
fn sanitize_and_validate_pipeline() {
    compile_ok(
        r#"<?php
$raw = '  User@Example.COM  ';
$sanitized = strtolower(trim(filter_var($raw, FILTER_SANITIZE_EMAIL)));
$valid = filter_var($sanitized, FILTER_VALIDATE_EMAIL);
echo $valid !== false ? $valid : 'invalid';
"#,
    );
}
