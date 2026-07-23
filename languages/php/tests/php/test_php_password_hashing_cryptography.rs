use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Password Hashing & Cryptography — password_hash, password_verify, hash, hash_hmac, hash_equals, random_bytes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_password_hash_and_verify_bcrypt() {
    let out = run_prints(
        r#"<?php
$pwd = "secret123Pass";
$hash = password_hash($pwd, PASSWORD_DEFAULT);

echo password_verify("secret123Pass", $hash) ? "VERIFIED" : "FAIL";
echo " | ";
echo password_verify("wrong_password", $hash) ? "VERIFIED" : "FAIL";
"#,
    );
    assert_eq!(out, vec!["VERIFIED | FAIL"]);
}

#[test]
fn test_php_password_needs_rehash_options() {
    let out = run_prints(
        r#"<?php
$hash = password_hash("secret", PASSWORD_BCRYPT, ["cost" => 4]);
$needs = password_needs_rehash($hash, PASSWORD_BCRYPT, ["cost" => 10]);
echo $needs ? "NEEDS_REHASH" : "OK";
"#,
    );
    assert_eq!(out, vec!["NEEDS_REHASH"]);
}

#[test]
fn test_php_hash_hmac_sha256() {
    let out = run_prints(
        r#"<?php
$message = "payload_data";
$key = "secret_key_123";
$hmac = hash_hmac("sha256", $message, $key);
echo strlen($hmac) === 64 ? "HMAC_LENGTH_64" : "INVALID";
"#,
    );
    assert_eq!(out, vec!["HMAC_LENGTH_64"]);
}

#[test]
fn test_php_hash_equals_timing_attack_prevention() {
    let out = run_prints(
        r#"<?php
$h1 = hash("sha256", "token123");
$h2 = hash("sha256", "token123");
$h3 = hash("sha256", "token456");

echo hash_equals($h1, $h2) ? "1" : "0";
echo hash_equals($h1, $h3) ? "1" : "0";
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_php_random_bytes_and_random_int() {
    let out = run_prints(
        r#"<?php
$bytes = random_bytes(16);
$randNum = random_int(1, 100);

echo (strlen($bytes) === 16 && $randNum >= 1 && $randNum <= 100) ? "CSPRNG_OK" : "FAIL";
"#,
    );
    assert_eq!(out, vec!["CSPRNG_OK"]);
}

#[test]
fn test_php_password_get_info_algorithm_cost() {
    compile_ok(
        r#"<?php
$hash = password_hash("pwd", PASSWORD_BCRYPT, ["cost" => 5]);
$info = password_get_info($hash);
echo "Algo=" . $info["algoName"] . " Cost=" . $info["options"]["cost"];
"#,
    );
}

#[test]
fn test_php_hash_algos_list_availability() {
    compile_ok(
        r#"<?php
$algos = hash_algos();
echo in_array("sha256", $algos) && in_array("md5", $algos) ? "ALGOS_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_hash_hkdf_key_derivation() {
    compile_ok(
        r#"<?php
$ikm = "input_key_material";
$derived = hash_hkdf("sha256", $ikm, 32, "info_label");
echo strlen($derived) === 32 ? "HKDF_32_BYTES" : "FAIL";
"#,
    );
}

#[test]
fn test_php_hash_pbkdf2_key_derivation() {
    compile_ok(
        r#"<?php
$derived = hash_pbkdf2("sha256", "password", "salt", 1000, 32);
echo strlen($derived) === 64 ? "HEX_LEN_64" : "FAIL";
"#,
    );
}

#[test]
fn test_php_crypt_des_md5_sha512_hashes() {
    compile_ok(
        r#"<?php
$hashed = crypt("my_password", '$6$rounds=5000$usesomesalt$');
echo strlen($hashed) > 0 ? "CRYPT_OK" : "FAIL";
"#,
    );
}
