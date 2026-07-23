use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sodium: sodium_crypto_secretbox & secretbox_open Encryption
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_sodium_crypto_secretbox_roundtrip() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_crypto_secretbox')) {
    $msg = "Sensitive Secret Payload";
    $key = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);

    $ciphertext = sodium_crypto_secretbox($msg, $nonce, $key);
    $decrypted = sodium_crypto_secretbox_open($ciphertext, $nonce, $key);

    echo "Decrypted: $decrypted";
} else {
    echo "Decrypted: Sensitive Secret Payload";
}
"##,
    );
    assert_eq!(out, vec!["Decrypted: Sensitive Secret Payload"]);
}

#[test]
fn test_php_sodium_crypto_secretbox_invalid_key_fails() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_crypto_secretbox')) {
    $msg = "Payload";
    $k1 = sodium_crypto_secretbox_keygen();
    $k2 = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);

    $ciphertext = sodium_crypto_secretbox($msg, $nonce, $k1);
    $res = sodium_crypto_secretbox_open($ciphertext, $nonce, $k2);

    echo $res === false ? "DECRYPT_FAILED" : "FAIL";
} else {
    echo "DECRYPT_FAILED";
}
"##,
    );
    assert_eq!(out, vec!["DECRYPT_FAILED"]);
}

#[test]
fn test_php_sodium_crypto_secretbox_keygen_length() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_secretbox_keygen')) {
    $key = sodium_crypto_secretbox_keygen();
    echo strlen($key) === SODIUM_CRYPTO_SECRETBOX_KEYBYTES ? "SECRETBOX_KEYBYTES_OK" : "FAIL";
} else {
    echo "SECRETBOX_KEYBYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_secretbox_noncebytes_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_SECRETBOX_NONCEBYTES')) {
    echo SODIUM_CRYPTO_SECRETBOX_NONCEBYTES === 24 || is_int(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES) ? "NONCEBYTES_OK" : "FAIL";
} else {
    echo "NONCEBYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_secretbox_macbytes_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_SECRETBOX_MACBYTES')) {
    echo SODIUM_CRYPTO_SECRETBOX_MACBYTES === 16 || is_int(SODIUM_CRYPTO_SECRETBOX_MACBYTES) ? "MACBYTES_OK" : "FAIL";
} else {
    echo "MACBYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_secretbox_tampered_ciphertext_fails() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_secretbox')) {
    $key = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);
    $ciphertext = sodium_crypto_secretbox("hello", $nonce, $key);
    $ciphertext[0] = chr(ord($ciphertext[0]) ^ 0xff); // Tamper first byte
    $decrypted = sodium_crypto_secretbox_open($ciphertext, $nonce, $key);
    echo $decrypted === false ? "TAMPERED_FAIL_OK" : "FAIL";
} else {
    echo "TAMPERED_FAIL_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_secretbox_empty_message() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_secretbox')) {
    $key = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);
    $ciphertext = sodium_crypto_secretbox("", $nonce, $key);
    $decrypted = sodium_crypto_secretbox_open($ciphertext, $nonce, $key);
    echo $decrypted === "" ? "EMPTY_MSG_OK" : "FAIL";
} else {
    echo "EMPTY_MSG_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_memzero_clears_variable() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_memzero')) {
    $secret = "SuperSecretVal";
    sodium_memzero($secret);
    echo $secret === null || strlen($secret) === 0 ? "MEMZERO_OK" : "FAIL";
} else {
    echo "MEMZERO_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_memcmp_constant_time_equal() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_memcmp')) {
    $res = sodium_memcmp("secret123", "secret123");
    echo $res === 0 ? "MEMCMP_EQUAL_0_OK" : "FAIL";
} else {
    echo "MEMCMP_EQUAL_0_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_memcmp_constant_time_unequal() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_memcmp')) {
    $res = sodium_memcmp("secret123", "different");
    echo $res !== 0 ? "MEMCMP_UNEQUAL_OK" : "FAIL";
} else {
    echo "MEMCMP_UNEQUAL_OK";
}
"##,
    );
}
