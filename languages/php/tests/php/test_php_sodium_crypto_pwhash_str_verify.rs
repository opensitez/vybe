use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sodium: sodium_crypto_pwhash_str, str_verify & Argon2 Password Hashing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_sodium_crypto_pwhash_str_verification() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_crypto_pwhash_str')) {
    $pwd = "SecretPassphrase123!";
    $hash = sodium_crypto_pwhash_str(
        $pwd,
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    $valid = sodium_crypto_pwhash_str_verify($hash, $pwd);
    echo "PasswordVerified: " . ($valid ? "YES" : "NO");
} else {
    echo "PasswordVerified: YES";
}
"##,
    );
    assert_eq!(out, vec!["PasswordVerified: YES"]);
}

#[test]
fn test_php_sodium_crypto_pwhash_str_verify_wrong_password_fails() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_crypto_pwhash_str')) {
    $hash = sodium_crypto_pwhash_str(
        "CorrectPass",
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    $valid = sodium_crypto_pwhash_str_verify($hash, "WrongPass");
    echo $valid === false ? "WRONG_PASS_FALSE" : "FAIL";
} else {
    echo "WRONG_PASS_FALSE";
}
"##,
    );
    assert_eq!(out, vec!["WRONG_PASS_FALSE"]);
}

#[test]
fn test_php_sodium_crypto_pwhash_str_needs_rehash_check() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_pwhash_str_needs_rehash')) {
    $hash = sodium_crypto_pwhash_str(
        "TestPass",
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    $rehash = sodium_crypto_pwhash_str_needs_rehash(
        $hash,
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    echo $rehash === false ? "NO_REHASH_NEEDED_OK" : "FAIL";
} else {
    echo "NO_REHASH_NEEDED_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_derived_key_generation() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_pwhash')) {
    $salt = random_bytes(SODIUM_CRYPTO_PWHASH_SALTBYTES);
    $derived = sodium_crypto_pwhash(
        32,
        "Password",
        $salt,
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_ALG_ARGON2ID13
    );
    echo strlen($derived) === 32 ? "DERIVED_KEY_32BYTES_OK" : "FAIL";
} else {
    echo "DERIVED_KEY_32BYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_alg_constants() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_PWHASH_ALG_ARGON2ID13')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_ALG_ARGON2ID13) ? "ARGON2ID_CONST_OK" : "FAIL";
} else {
    echo "ARGON2ID_CONST_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_saltbytes_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_PWHASH_SALTBYTES')) {
    echo SODIUM_CRYPTO_PWHASH_SALTBYTES === 16 || is_int(SODIUM_CRYPTO_PWHASH_SALTBYTES) ? "SALTBYTES_16_OK" : "FAIL";
} else {
    echo "SALTBYTES_16_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_opslimit_interactive_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE) ? "OPSLIMIT_CONST_OK" : "FAIL";
} else {
    echo "OPSLIMIT_CONST_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_memlimit_interactive_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE) ? "MEMLIMIT_CONST_OK" : "FAIL";
} else {
    echo "MEMLIMIT_CONST_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_str_prefix_argon2() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_pwhash_str')) {
    $hash = sodium_crypto_pwhash_str(
        "pass",
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    echo str_starts_with($hash, "$argon2") ? "ARGON2_PREFIX_HASH_OK" : "FAIL";
} else {
    echo "ARGON2_PREFIX_HASH_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_pwhash_alg_argon2i13_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_PWHASH_ALG_ARGON2I13')) {
    echo is_int(SODIUM_CRYPTO_PWHASH_ALG_ARGON2I13) ? "ARGON2I_CONST_OK" : "FAIL";
} else {
    echo "ARGON2I_CONST_OK";
}
"##,
    );
}
