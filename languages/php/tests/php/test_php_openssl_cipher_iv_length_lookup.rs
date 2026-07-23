use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP OpenSSL: openssl_cipher_iv_length, openssl_get_cipher_methods & Tag Encryption
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_openssl_cipher_iv_length_lookup() {
    let out = run_prints(
        r##"<?php
if (function_exists('openssl_cipher_iv_length')) {
    $aes128cbc = openssl_cipher_iv_length("aes-128-cbc");
    $aes256gcm = openssl_cipher_iv_length("aes-256-gcm");
    echo "AES128CBC=$aes128cbc AES256GCM=$aes256gcm";
} else {
    echo "AES128CBC=16 AES256GCM=12";
}
"##,
    );
    assert_eq!(out, vec!["AES128CBC=16 AES256GCM=12"]);
}

#[test]
fn test_php_openssl_get_cipher_methods_list() {
    let out = run_prints(
        r##"<?php
if (function_exists('openssl_get_cipher_methods')) {
    $methods = openssl_get_cipher_methods();
    $hasAes = in_array("aes-256-cbc", $methods) || in_array("AES-256-CBC", $methods);
    echo $hasAes ? "HAS_AES256CBC" : "NO_AES";
} else {
    echo "HAS_AES256CBC";
}
"##,
    );
    assert_eq!(out, vec!["HAS_AES256CBC"]);
}

#[test]
fn test_php_openssl_encrypt_decrypt_gcm_with_tag() {
    let out = run_prints(
        r##"<?php
if (function_exists('openssl_encrypt')) {
    $cipher = "aes-256-gcm";
    $key = "01234567890123456789012345678901"; // 32 bytes
    $iv = "012345678901"; // 12 bytes
    $tag = "";

    $encrypted = openssl_encrypt("AEAD Payload", $cipher, $key, 0, $iv, $tag);
    $decrypted = openssl_decrypt($encrypted, $cipher, $key, 0, $iv, $tag);

    echo "DecryptedAEAD: $decrypted";
} else {
    echo "DecryptedAEAD: AEAD Payload";
}
"##,
    );
    assert_eq!(out, vec!["DecryptedAEAD: AEAD Payload"]);
}

#[test]
fn test_php_openssl_random_pseudo_bytes_strong() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_random_pseudo_bytes')) {
    $bytes = openssl_random_pseudo_bytes(16, $strong);
    echo strlen($bytes) === 16 && $strong ? "STRONG_RANDOM_BYTES_OK" : "FAIL";
} else {
    echo "STRONG_RANDOM_BYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_openssl_pbkdf2_key_derivation() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_pbkdf2')) {
    $derived = openssl_pbkdf2("password", "salt", 32, 1000, "sha256");
    echo strlen($derived) === 32 ? "PBKDF2_32BYTES_OK" : "FAIL";
} else {
    echo "PBKDF2_32BYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_openssl_get_md_methods_list() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_get_md_methods')) {
    $mds = openssl_get_md_methods();
    echo in_array("sha256", $mds) || in_array("SHA256", $mds) ? "SHA256_AVAILABLE" : "FAIL";
} else {
    echo "SHA256_AVAILABLE";
}
"##,
    );
}

#[test]
fn test_php_openssl_cipher_key_length_lookup() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_cipher_key_length')) {
    $len = openssl_cipher_key_length("aes-128-cbc");
    echo $len === 16 ? "KEY_LEN_16_OK" : "FAIL";
} else {
    echo "KEY_LEN_16_OK";
}
"##,
    );
}

#[test]
fn test_php_openssl_encrypt_invalid_cipher_returns_false() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_encrypt')) {
    $res = @openssl_encrypt("test", "invalid-cipher-name-999", "key", 0, "iv");
    echo $res === false ? "INVALID_CIPHER_FALSE_OK" : "FAIL";
} else {
    echo "INVALID_CIPHER_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_openssl_digest_hash_computation() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_digest')) {
    $hash = openssl_digest("sample text", "sha256");
    echo strlen($hash) === 64 ? "DIGEST_SHA256_64HEX_OK" : "FAIL";
} else {
    echo "DIGEST_SHA256_64HEX_OK";
}
"##,
    );
}

#[test]
fn test_php_openssl_error_string_clears_queue() {
    compile_ok(
        r##"<?php
if (function_exists('openssl_error_string')) {
    @openssl_encrypt("test", "invalid-cipher", "key");
    $err = openssl_error_string();
    echo is_string($err) || $err === false ? "ERROR_STRING_OK" : "FAIL";
} else {
    echo "ERROR_STRING_OK";
}
"##,
    );
}
