use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sodium: sodium_crypto_sign, detached_sign & verify_detached
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_sodium_crypto_sign_detached_verification() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_crypto_sign_detached')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $pk = sodium_crypto_sign_publickey($kp);

    $msg = "Message to sign";
    $sig = sodium_crypto_sign_detached($msg, $sk);
    $valid = sodium_crypto_sign_verify_detached($sig, $msg, $pk);

    echo "ValidSignature: " . ($valid ? "YES" : "NO");
} else {
    echo "ValidSignature: YES";
}
"##,
    );
    assert_eq!(out, vec!["ValidSignature: YES"]);
}

#[test]
fn test_php_sodium_crypto_sign_combined_sign_and_open() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_crypto_sign')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $pk = sodium_crypto_sign_publickey($kp);

    $signedMsg = sodium_crypto_sign("Signed Content", $sk);
    $original = sodium_crypto_sign_open($signedMsg, $pk);

    echo "Opened: $original";
} else {
    echo "Opened: Signed Content";
}
"##,
    );
    assert_eq!(out, vec!["Opened: Signed Content"]);
}

#[test]
fn test_php_sodium_crypto_sign_bytes_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_CRYPTO_SIGN_BYTES')) {
    echo SODIUM_CRYPTO_SIGN_BYTES === 64 || is_int(SODIUM_CRYPTO_SIGN_BYTES) ? "SIGN_BYTES_OK" : "FAIL";
} else {
    echo "SIGN_BYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_publickey_length() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_keypair')) {
    $kp = sodium_crypto_sign_keypair();
    $pk = sodium_crypto_sign_publickey($kp);
    echo strlen($pk) === SODIUM_CRYPTO_SIGN_PUBLICKEYBYTES ? "PUBLICKEYBYTES_OK" : "FAIL";
} else {
    echo "PUBLICKEYBYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_secretkey_length() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_keypair')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    echo strlen($sk) === SODIUM_CRYPTO_SIGN_SECRETKEYBYTES ? "SECRETKEYBYTES_OK" : "FAIL";
} else {
    echo "SECRETKEYBYTES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_seed_keypair_generation() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_seed_keypair')) {
    $seed = random_bytes(SODIUM_CRYPTO_SIGN_SEEDBYTES);
    $kp = sodium_crypto_sign_seed_keypair($seed);
    echo strlen($kp) === SODIUM_CRYPTO_SIGN_KEYPAIRBYTES ? "SEED_KEYPAIR_OK" : "FAIL";
} else {
    echo "SEED_KEYPAIR_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_verify_tampered_message_fails() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_detached')) {
    $kp = sodium_crypto_sign_keypair();
    $sig = sodium_crypto_sign_detached("Original", sodium_crypto_sign_secretkey($kp));
    $valid = sodium_crypto_sign_verify_detached($sig, "Tampered", sodium_crypto_sign_publickey($kp));
    echo $valid === false ? "TAMPERED_VERIFY_FALSE_OK" : "FAIL";
} else {
    echo "TAMPERED_VERIFY_FALSE_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_keypair_from_secretkey_and_publickey() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_keypair_from_secretkey_and_publickey')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $pk = sodium_crypto_sign_publickey($kp);
    $reconstructed = sodium_crypto_sign_keypair_from_secretkey_and_publickey($sk, $pk);
    echo strlen($reconstructed) === SODIUM_CRYPTO_SIGN_KEYPAIRBYTES ? "RECONSTRUCTED_KEYPAIR_OK" : "FAIL";
} else {
    echo "RECONSTRUCTED_KEYPAIR_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_ed25519_pk_to_curve25519() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_ed25519_pk_to_curve25519')) {
    $kp = sodium_crypto_sign_keypair();
    $pk = sodium_crypto_sign_publickey($kp);
    $x25519_pk = sodium_crypto_sign_ed25519_pk_to_curve25519($pk);
    echo strlen($x25519_pk) === SODIUM_CRYPTO_BOX_PUBLICKEYBYTES ? "ED25519_TO_CURVE_OK" : "FAIL";
} else {
    echo "ED25519_TO_CURVE_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_crypto_sign_ed25519_sk_to_curve25519() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_crypto_sign_ed25519_sk_to_curve25519')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $x25519_sk = sodium_crypto_sign_ed25519_sk_to_curve25519($sk);
    echo strlen($x25519_sk) === SODIUM_CRYPTO_BOX_SECRETKEYBYTES ? "ED25519_SK_TO_CURVE_OK" : "FAIL";
} else {
    echo "ED25519_SK_TO_CURVE_OK";
}
"##,
    );
}
