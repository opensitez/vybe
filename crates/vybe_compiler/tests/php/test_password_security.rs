use super::helpers::compile_ok;

// ── password_hash / password_verify ──────────────────────────

#[test] fn password_hash_default() {
    compile_ok(r#"<?php
$hash = password_hash('secret123', PASSWORD_DEFAULT);
echo is_string($hash) ? 'hashed' : 'fail';
echo strlen($hash) > 20 ? ':long enough' : ':too short';
"#);
}

#[test] fn password_hash_bcrypt() {
    compile_ok(r#"<?php
$hash = password_hash('mypassword', PASSWORD_BCRYPT);
echo str_starts_with($hash, '$2y$') ? 'bcrypt hash' : 'not bcrypt';
"#);
}

#[test] fn password_hash_argon2i() {
    compile_ok(r#"<?php
if (!defined('PASSWORD_ARGON2I')) {
    echo 'argon2i not available';
} else {
    $hash = password_hash('mypassword', PASSWORD_ARGON2I);
    echo str_contains($hash, 'argon2i') ? 'argon2i hash' : 'different algo';
}
"#);
}

#[test] fn password_hash_argon2id() {
    compile_ok(r#"<?php
if (!defined('PASSWORD_ARGON2ID')) {
    echo 'argon2id not available';
} else {
    $hash = password_hash('mypassword', PASSWORD_ARGON2ID);
    echo str_contains($hash, 'argon2id') ? 'argon2id hash' : 'different algo';
}
"#);
}

#[test] fn password_verify_correct() {
    compile_ok(r#"<?php
$password = 'correcthorsebatterystaple';
$hash = password_hash($password, PASSWORD_DEFAULT);
echo password_verify($password, $hash) ? 'correct' : 'wrong';
"#);
}

#[test] fn password_verify_wrong() {
    compile_ok(r#"<?php
$hash = password_hash('rightpassword', PASSWORD_DEFAULT);
echo password_verify('wrongpassword', $hash) ? 'matches' : 'no match';
echo password_verify('',              $hash) ? 'matches' : 'no match';
"#);
}

#[test] fn password_verify_tampered_hash() {
    compile_ok(r#"<?php
$hash = password_hash('secret', PASSWORD_DEFAULT);
$tampered = substr($hash, 0, 10) . 'XXXX' . substr($hash, 14);
echo password_verify('secret', $tampered) ? 'verified' : 'failed';
"#);
}

// ── password_needs_rehash ─────────────────────────────────────

#[test] fn password_needs_rehash_same_algo() {
    compile_ok(r#"<?php
$hash = password_hash('pass', PASSWORD_BCRYPT, ['cost' => 10]);
$needs = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 10]);
echo $needs ? 'needs rehash' : 'ok';
"#);
}

#[test] fn password_needs_rehash_higher_cost() {
    compile_ok(r#"<?php
$hash = password_hash('pass', PASSWORD_BCRYPT, ['cost' => 10]);
$needs = password_needs_rehash($hash, PASSWORD_BCRYPT, ['cost' => 12]);
echo $needs ? 'needs rehash' : 'ok';
"#);
}

#[test] fn password_needs_rehash_different_algo() {
    compile_ok(r#"<?php
$hash = password_hash('pass', PASSWORD_BCRYPT);
$needs = password_needs_rehash($hash, PASSWORD_DEFAULT);
// If DEFAULT changed, might need rehash
echo is_bool($needs) ? 'bool result' : 'not bool';
"#);
}

// ── password_get_info ─────────────────────────────────────────

#[test] fn password_get_info_bcrypt() {
    compile_ok(r#"<?php
$hash = password_hash('test', PASSWORD_BCRYPT);
$info = password_get_info($hash);
echo isset($info['algo']) ? 'has algo' : 'no algo';
echo isset($info['algoName']) ? ':has algoName' : ':no algoName';
echo isset($info['options']) ? ':has options' : ':no options';
"#);
}

#[test] fn password_get_info_unknown() {
    compile_ok(r#"<?php
$info = password_get_info('not-a-hash');
echo $info['algo'] === 0 ? 'unknown algo' : 'known algo';
"#);
}

// ── bcrypt cost options ───────────────────────────────────────

#[test] fn password_bcrypt_cost() {
    compile_ok(r#"<?php
$hash = password_hash('test', PASSWORD_BCRYPT, ['cost' => 4]);
$info = password_get_info($hash);
echo $info['options']['cost'] === 4 ? 'cost 4' : 'different cost';
"#);
}

// ── hash / hash_hmac ──────────────────────────────────────────

#[test] fn hash_sha256() {
    compile_ok(r#"<?php
$h = hash('sha256', 'hello');
echo strlen($h) === 64 ? 'sha256 length ok' : 'wrong length';
echo ctype_xdigit($h) ? ':hex chars' : ':not hex';
"#);
}

#[test] fn hash_sha512() {
    compile_ok(r#"<?php
$h = hash('sha512', 'hello');
echo strlen($h) === 128 ? 'sha512 length ok' : 'wrong length';
"#);
}

#[test] fn hash_md5_function() {
    compile_ok(r#"<?php
$h = hash('md5', 'hello');
echo strlen($h) === 32 ? 'md5 length ok' : 'wrong length';
echo $h === md5('hello') ? ':matches md5()' : ':different';
"#);
}

#[test] fn hash_algorithms_available() {
    compile_ok(r#"<?php
$algos = hash_algos();
echo in_array('sha256', $algos) ? 'sha256 available' : 'missing sha256';
echo in_array('sha512', $algos) ? ':sha512 available' : ':missing sha512';
echo in_array('md5',    $algos) ? ':md5 available'    : ':missing md5';
"#);
}

#[test] fn hash_binary_output() {
    compile_ok(r#"<?php
$hex = hash('sha256', 'test', false);
$bin = hash('sha256', 'test', true);
echo strlen($hex) === 64 ? 'hex 64 chars' : 'wrong';
echo strlen($bin) === 32 ? ':bin 32 bytes' : ':wrong';
"#);
}

#[test] fn hash_hmac_basic() {
    compile_ok(r#"<?php
$key = 'secret-key';
$msg = 'message to sign';
$hmac = hash_hmac('sha256', $msg, $key);
echo strlen($hmac) === 64 ? 'hmac length ok' : 'wrong length';
"#);
}

#[test] fn hash_hmac_verify() {
    compile_ok(r#"<?php
$key  = 'my-secret';
$data = 'payload';
$sig  = hash_hmac('sha256', $data, $key);
$expected = hash_hmac('sha256', $data, $key);
echo hash_equals($sig, $expected) ? 'valid' : 'tampered';
"#);
}

#[test] fn hash_hmac_tampered() {
    compile_ok(r#"<?php
$key  = 'secret';
$data = 'original';
$sig  = hash_hmac('sha256', $data, $key);
$other = hash_hmac('sha256', 'tampered', $key);
echo hash_equals($sig, $other) ? 'matches' : 'no match';
"#);
}

// ── hash_equals (timing-safe comparison) ─────────────────────

#[test] fn hash_equals_same() {
    compile_ok(r#"<?php
$a = hash('sha256', 'hello');
$b = hash('sha256', 'hello');
echo hash_equals($a, $b) ? 'equal' : 'not equal';
"#);
}

#[test] fn hash_equals_different() {
    compile_ok(r#"<?php
$a = hash('sha256', 'hello');
$b = hash('sha256', 'world');
echo hash_equals($a, $b) ? 'equal' : 'not equal';
"#);
}

// ── Practical security patterns ───────────────────────────────

#[test] fn secure_token_generation() {
    compile_ok(r#"<?php
function generateToken(int $bytes = 32): string {
    return bin2hex(random_bytes($bytes));
}
$token = generateToken();
echo strlen($token) === 64 ? 'token length ok' : 'wrong length';
echo ctype_xdigit($token) ? ':hex' : ':not hex';
"#);
}

#[test] fn secure_password_workflow() {
    compile_ok(r#"<?php
// Registration
$plaintext = 'user_password_123';
$stored = password_hash($plaintext, PASSWORD_DEFAULT);
// Login verification
$attempt = 'user_password_123';
$valid = password_verify($attempt, $stored);
// Check if needs upgrade
$upgrade = password_needs_rehash($stored, PASSWORD_DEFAULT);
echo $valid ? 'authenticated' : 'denied';
echo !$upgrade ? ':hash current' : ':needs upgrade';
"#);
}

#[test] fn api_key_hashing() {
    compile_ok(r#"<?php
function hashApiKey(string $key): string {
    return hash('sha256', $key);
}
function verifyApiKey(string $key, string $stored): bool {
    return hash_equals($stored, hashApiKey($key));
}
$raw = bin2hex(random_bytes(16));
$stored = hashApiKey($raw);
echo verifyApiKey($raw, $stored)         ? 'valid' : 'invalid';
echo verifyApiKey('wrong-key', $stored)  ? 'valid' : ':invalid';
"#);
}
