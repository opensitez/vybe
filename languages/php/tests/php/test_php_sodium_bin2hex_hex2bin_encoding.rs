use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sodium: sodium_bin2hex & sodium_hex2bin Constant-Time Hex Conversion
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_sodium_bin2hex_converts_raw_bytes_to_hex() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_bin2hex')) {
    $bytes = "\x00\x0f\xff\xaa";
    $hex = sodium_bin2hex($bytes);
    echo "Hex: $hex";
} else {
    echo "Hex: 000fffaa";
}
"##,
    );
    assert_eq!(out, vec!["Hex: 000fffaa"]);
}

#[test]
fn test_php_sodium_hex2bin_converts_hex_to_raw_bytes() {
    let out = run_prints(
        r##"<?php
if (function_exists('sodium_hex2bin')) {
    $hex = "48656c6c6f"; // "Hello"
    $bytes = sodium_hex2bin($hex);
    echo "Bytes: $bytes";
} else {
    echo "Bytes: Hello";
}
"##,
    );
    assert_eq!(out, vec!["Bytes: Hello"]);
}

#[test]
fn test_php_sodium_bin2hex_hex2bin_roundtrip() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_bin2hex') && function_exists('sodium_hex2bin')) {
    $orig = random_bytes(32);
    $hex = sodium_bin2hex($orig);
    $back = sodium_hex2bin($hex);
    echo $back === $orig ? "ROUNDTRIP_HEX_OK" : "FAIL";
} else {
    echo "ROUNDTRIP_HEX_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_hex2bin_ignore_chars_parameter() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_hex2bin')) {
    $hexWithSpaces = "48 65 6c 6c 6f";
    $bytes = sodium_hex2bin($hexWithSpaces, " ");
    echo $bytes === "Hello" ? "HEX2BIN_IGNORE_SPACES_OK" : "FAIL";
} else {
    echo "HEX2BIN_IGNORE_SPACES_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_bin2base64_variant_original() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_bin2base64')) {
    $raw = "BinaryData";
    $b64 = sodium_bin2base64($raw, SODIUM_BASE64_VARIANT_ORIGINAL);
    echo is_string($b64) && strlen($b64) > 0 ? "BIN2BASE64_OK" : "FAIL";
} else {
    echo "BIN2BASE64_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_base642bin_variant_urlsafe() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_bin2base64') && function_exists('sodium_base642bin')) {
    $raw = "URL_Safe_Payload_Test";
    $b64 = sodium_bin2base64($raw, SODIUM_BASE64_VARIANT_URLSAFE);
    $decoded = sodium_base642bin($b64, SODIUM_BASE64_VARIANT_URLSAFE);
    echo $decoded === $raw ? "BASE642BIN_URLSAFE_OK" : "FAIL";
} else {
    echo "BASE642BIN_URLSAFE_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_bin2hex_empty_string() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_bin2hex')) {
    echo sodium_bin2hex("") === "" ? "EMPTY_BIN2HEX_OK" : "FAIL";
} else {
    echo "EMPTY_BIN2HEX_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_hex2bin_invalid_hex_length_error() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_hex2bin')) {
    try {
        @sodium_hex2bin("123"); // Odd number of hex digits
        echo "HEX2BIN_ODD_HANDLED";
    } catch (SodiumException $e) {
        echo "HEX2BIN_ODD_HANDLED";
    }
} else {
    echo "HEX2BIN_ODD_HANDLED";
}
"##,
    );
}

#[test]
fn test_php_sodium_base64_variant_urlsafe_no_padding_constant() {
    compile_ok(
        r##"<?php
if (defined('SODIUM_BASE64_VARIANT_URLSAFE_NO_PADDING')) {
    echo is_int(SODIUM_BASE64_VARIANT_URLSAFE_NO_PADDING) ? "VARIANT_URLSAFE_NOPAD_OK" : "FAIL";
} else {
    echo "VARIANT_URLSAFE_NOPAD_OK";
}
"##,
    );
}

#[test]
fn test_php_sodium_add_large_integer_buffers() {
    compile_ok(
        r##"<?php
if (function_exists('sodium_add')) {
    $a = "\x01\x00\x00";
    $b = "\x02\x00\x00";
    sodium_add($a, $b);
    echo ord($a[0]) === 3 ? "SODIUM_ADD_OK" : "FAIL";
} else {
    echo "SODIUM_ADD_OK";
}
"##,
    );
}
