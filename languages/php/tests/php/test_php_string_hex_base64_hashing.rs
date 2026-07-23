use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: String Encoding & Cryptographic Hashing — base64_encode, base64_decode, hex2bin, bin2hex, crc32, md5, sha1, hash
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_base64_encode_and_decode_roundtrip() {
    let out = run_prints(
        r#"<?php
$original = "Hello World! Binary \x00\x01\x02";
$encoded = base64_encode($original);
$decoded = base64_decode($encoded);

echo ($decoded === $original ? "BASE64_ROUNDTRIP_OK" : "FAIL");
"#,
    );
    assert_eq!(out, vec!["BASE64_ROUNDTRIP_OK"]);
}

#[test]
fn test_php_bin2hex_and_hex2bin_conversions() {
    let out = run_prints(
        r#"<?php
$binary = "\x48\x65\x6c\x6c\x6f"; // "Hello"
$hex = bin2hex($binary);
$restored = hex2bin($hex);

echo "$hex | $restored";
"#,
    );
    assert_eq!(out, vec!["48656c6c6f | Hello"]);
}

#[test]
fn test_php_crc32_checksum_calculation() {
    let out = run_prints(
        r#"<?php
$checksum = crc32("The quick brown fox jumps over the lazy dog");
echo is_int($checksum) ? "CRC32_INT_OK" : "FAIL";
"#,
    );
    assert_eq!(out, vec!["CRC32_INT_OK"]);
}

#[test]
fn test_php_md5_and_sha1_hex_digests() {
    let out = run_prints(
        r#"<?php
$str = "apple";
$md5 = md5($str);
$sha1 = sha1($str);

echo "md5_len=" . strlen($md5) . " sha1_len=" . strlen($sha1);
"#,
    );
    assert_eq!(out, vec!["md5_len=32 sha1_len=40"]);
}

#[test]
fn test_php_hash_raw_output_binary_mode() {
    compile_ok(
        r#"<?php
$rawHash = hash("sha256", "data", binary: true);
echo strlen($rawHash) === 32 ? "RAW_32_BYTES" : "FAIL";
"#,
    );
}

#[test]
fn test_php_hash_init_update_final_incremental() {
    compile_ok(
        r#"<?php
$ctx = hash_init("sha256");
hash_update($ctx, "Part 1 ");
hash_update($ctx, "Part 2");
$digest = hash_final($ctx);

echo strlen($digest) === 64 ? "INCREMENTAL_HASH_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_base64_decode_strict_mode() {
    compile_ok(
        r#"<?php
$invalidBase64 = "===invalid===";
$res = base64_decode($invalidBase64, strict: true);
echo $res === false ? "STRICT_DECODE_FAILED" : "DECODED";
"#,
    );
}

#[test]
fn test_php_hash_copy_context() {
    compile_ok(
        r#"<?php
$ctx1 = hash_init("md5");
hash_update($ctx1, "hello");
$ctx2 = hash_copy($ctx1);
hash_update($ctx1, " world");
hash_update($ctx2, " php");

echo hash_final($ctx1) !== hash_final($ctx2) ? "DIFFERENT_HASHES" : "SAME";
"#,
    );
}

#[test]
fn test_php_hash_hmac_file_checksum() {
    compile_ok(
        r#"<?php
$tmp = tempnam(sys_get_temp_dir(), "hash_file_");
file_put_contents($tmp, "file content for hmac");
$hmac = hash_hmac_file("sha256", $tmp, "secret_key");
unlink($tmp);

echo strlen($hmac) === 64 ? "HMAC_FILE_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_hash_file_digest() {
    compile_ok(
        r#"<?php
$tmp = tempnam(sys_get_temp_dir(), "hash_digest_");
file_put_contents($tmp, "content");
$digest = hash_file("md5", $tmp);
unlink($tmp);

echo strlen($digest) === 32 ? "HASH_FILE_OK" : "FAIL";
"#,
    );
}
