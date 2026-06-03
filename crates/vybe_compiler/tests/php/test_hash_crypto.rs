use super::helpers::compile_ok;

// ── hash() with various algorithms ──────────────────────────────

#[test]
fn hash_sha256_algo() {
    compile_ok(
        r#"<?php
$h = hash('sha256', 'hello world');
echo strlen($h) === 64 ? 'ok' : 'fail';
echo ctype_xdigit($h) ? ':hex' : ':not hex';
"#,
    );
}

#[test]
fn hash_sha512_algo() {
    compile_ok(
        r#"<?php
$h = hash('sha512', 'hello world');
echo strlen($h) === 128 ? 'ok' : 'fail';
echo ctype_xdigit($h) ? ':hex' : ':not hex';
"#,
    );
}

#[test]
fn hash_md5_via_hash() {
    compile_ok(
        r#"<?php
$h = hash('md5', 'hello');
echo strlen($h) === 32 ? 'ok' : 'fail';
echo $h === md5('hello') ? ':matches' : ':differs';
"#,
    );
}

#[test]
fn hash_sha1_via_hash() {
    compile_ok(
        r#"<?php
$h = hash('sha1', 'hello');
echo strlen($h) === 40 ? 'ok' : 'fail';
echo $h === sha1('hello') ? ':matches' : ':differs';
"#,
    );
}

#[test]
fn hash_raw_output() {
    compile_ok(
        r#"<?php
$hex = hash('sha256', 'test', false);
$bin = hash('sha256', 'test', true);
echo strlen($hex) === 64 ? 'hex ok' : 'hex fail';
echo strlen($bin) === 32 ? ':bin ok' : ':bin fail';
echo bin2hex($bin) === $hex ? ':round-trip ok' : ':round-trip fail';
"#,
    );
}

// ── hash_hmac ────────────────────────────────────────────────────

#[test]
fn hash_hmac_sha256() {
    compile_ok(
        r#"<?php
$key = 'secret';
$msg = 'data to sign';
$mac = hash_hmac('sha256', $msg, $key);
echo strlen($mac) === 64 ? 'ok' : 'fail';
echo ctype_xdigit($mac) ? ':hex' : ':not hex';
"#,
    );
}

// ── hash_algos ───────────────────────────────────────────────────

#[test]
fn hash_algos_list() {
    compile_ok(
        r#"<?php
$algos = hash_algos();
echo is_array($algos) ? 'array' : 'not array';
echo in_array('sha256', $algos) ? ':sha256' : ':no sha256';
echo in_array('md5', $algos) ? ':md5' : ':no md5';
"#,
    );
}

// ── hash_equals ──────────────────────────────────────────────────

#[test]
fn hash_equals_timing_safe() {
    compile_ok(
        r#"<?php
$a = hash('sha256', 'secret');
$b = hash('sha256', 'secret');
$c = hash('sha256', 'other');
echo hash_equals($a, $b) ? 'equal' : 'not equal';
echo hash_equals($a, $c) ? 'equal' : 'not equal';
"#,
    );
}

// ── md5_file / sha1_file ─────────────────────────────────────────

#[test]
fn md5_file_hash() {
    compile_ok(
        r#"<?php
$result = md5_file('/etc/hostname');
echo is_string($result) || $result === false ? 'ok' : 'fail';
if (is_string($result)) {
    echo strlen($result) === 32 ? ':len ok' : ':len fail';
}
"#,
    );
}

#[test]
fn sha1_file_hash() {
    compile_ok(
        r#"<?php
$result = sha1_file('/etc/hostname');
echo is_string($result) || $result === false ? 'ok' : 'fail';
if (is_string($result)) {
    echo strlen($result) === 40 ? ':len ok' : ':len fail';
}
"#,
    );
}

// ── password_hash / password_verify / password_needs_rehash / password_get_info

#[test]
fn password_hash_bcrypt_algo() {
    compile_ok(
        r#"<?php
$hash = password_hash('mypassword', PASSWORD_BCRYPT);
echo is_string($hash) ? 'hashed' : 'fail';
echo str_starts_with($hash, '$2y$') ? ':bcrypt' : ':other';
"#,
    );
}

#[test]
fn password_hash_argon2i_algo() {
    compile_ok(
        r#"<?php
if (defined('PASSWORD_ARGON2I')) {
    $hash = password_hash('mypassword', PASSWORD_ARGON2I);
    echo is_string($hash) ? 'hashed' : 'fail';
} else {
    echo 'argon2i unavailable';
}
"#,
    );
}

#[test]
fn password_verify_correct_and_wrong() {
    compile_ok(
        r#"<?php
$hash = password_hash('correct', PASSWORD_DEFAULT);
echo password_verify('correct', $hash) ? 'ok' : 'fail';
echo password_verify('wrong', $hash) ? 'ok' : ':rejected';
"#,
    );
}

#[test]
fn password_needs_rehash_check() {
    compile_ok(
        r#"<?php
$hash = password_hash('pass', PASSWORD_BCRYPT, ['cost' => 10]);
$same = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 10]);
$more = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 14]);
echo is_bool($same) ? 'bool' : 'not bool';
echo $more ? ':needs rehash' : ':no rehash';
"#,
    );
}

#[test]
fn password_get_info_structure() {
    compile_ok(
        r#"<?php
$hash = password_hash('test', PASSWORD_BCRYPT);
$info = password_get_info($hash);
echo array_key_exists('algo', $info) ? 'has algo' : 'no algo';
echo array_key_exists('algoName', $info) ? ':has name' : ':no name';
echo array_key_exists('options', $info) ? ':has opts' : ':no opts';
"#,
    );
}

// ── openssl_random_pseudo_bytes ───────────────────────────────────

#[test]
fn openssl_random_pseudo_bytes_length() {
    compile_ok(
        r#"<?php
$bytes = openssl_random_pseudo_bytes(16, $strong);
echo strlen($bytes) === 16 ? 'ok' : 'fail';
echo is_bool($strong) ? ':strong flag bool' : ':not bool';
"#,
    );
}

// ── random_bytes / random_int ─────────────────────────────────────

#[test]
fn random_bytes_secure() {
    compile_ok(
        r#"<?php
$bytes = random_bytes(32);
echo strlen($bytes) === 32 ? 'ok' : 'fail';
echo strlen(bin2hex($bytes)) === 64 ? ':hex ok' : ':hex fail';
"#,
    );
}

#[test]
fn random_int_range() {
    compile_ok(
        r#"<?php
$n = random_int(1, 100);
echo $n >= 1 && $n <= 100 ? 'in range' : 'out of range';
echo is_int($n) ? ':int' : ':not int';
"#,
    );
}

// ── crc32 ─────────────────────────────────────────────────────────

#[test]
fn crc32_hash() {
    compile_ok(
        r#"<?php
$checksum = crc32('hello world');
echo is_int($checksum) ? 'int' : 'not int';
echo crc32('hello world') === crc32('hello world') ? ':deterministic' : ':varies';
"#,
    );
}

// ── base64 round-trip with binary data ───────────────────────────

#[test]
fn base64_roundtrip_binary() {
    compile_ok(
        r#"<?php
$raw = random_bytes(24);
$encoded = base64_encode($raw);
$decoded = base64_decode($encoded);
echo $decoded === $raw ? 'roundtrip ok' : 'roundtrip fail';
echo ctype_print($encoded) ? ':printable' : ':not printable';
"#,
    );
}
