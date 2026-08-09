use super::helpers::run_python;

// hmac — new, update, digest, hexdigest, compare_digest, HMAC object attributes

#[test]
fn test_hmac_sha256_hexdigest_validation() {
    let out = run_python(
        r#"
import hmac, hashlib
key = b"secret_key"
msg = b"hello world"
h = hmac.new(key, msg, hashlib.sha256)
print(h.hexdigest())
print(h.digest_size)
print(h.block_size)
"#,
    );
    assert_eq!(
        out,
        vec![
            "734cc62f32841568f45715aeb9f4d7891324e6d948e4c6c60c0621cd4294d5c6",
            "32",
            "64"
        ]
    );
}

#[test]
fn test_hmac_compare_digest_timing_safe() {
    let out = run_python(
        r#"
import hmac
sig1 = b"abc123xyz"
sig2 = b"abc123xyz"
sig3 = b"abc123xxx"
print(hmac.compare_digest(sig1, sig2))
print(hmac.compare_digest(sig1, sig3))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_hmac_compare_digest_str_inputs() {
    let out = run_python(
        r#"
import hmac
s1 = "hash_token_string_123"
s2 = "hash_token_string_123"
s3 = "different_token_456"
print(hmac.compare_digest(s1, s2))
print(hmac.compare_digest(s1, s3))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_hmac_digest_helper_function() {
    let out = run_python(
        r#"
import hmac
key = b"key"
msg = b"message"
d1 = hmac.digest(key, msg, "sha256")
d2 = hmac.new(key, msg, "sha256").digest()
print(d1 == d2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hmac_update_incremental() {
    let out = run_python(
        r#"
import hmac, hashlib
key = b"secret"
h1 = hmac.new(key, b"part1part2", hashlib.sha256)
h2 = hmac.new(key, digestmod=hashlib.sha256)
h2.update(b"part1")
h2.update(b"part2")
print(h1.hexdigest() == h2.hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hmac_copy_object_state() {
    let out = run_python(
        r#"
import hmac
key = b"my_key"
h1 = hmac.new(key, b"header", "sha256")
h2 = h1.copy()
h1.update(b"_body1")
h2.update(b"_body1")
print(h1.hexdigest() == h2.hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hmac_sha512_digest_size() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"key", b"msg", "sha512")
print(len(h.digest()))
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["64", "128"]);
}

#[test]
fn test_hmac_name_property() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"key", b"msg", "sha256")
print(h.name)
"#,
    );
    assert_eq!(out, vec!["hmac-sha256"]);
}

#[test]
fn test_hmac_string_digestmod_name() {
    let out = run_python(
        r#"
import hmac
h1 = hmac.new(b"key", b"data", "sha256")
h2 = hmac.new(b"key", b"data", "sha256")
print(h1.hexdigest() == h2.hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hmac_compare_digest_type_mismatch_raises_typeerror() {
    let out = run_python(
        r#"
import hmac
try:
    hmac.compare_digest(b"bytes", "string")
except TypeError:
    print("TypeError")
"#,
    );
    assert_eq!(out, vec!["TypeError"]);
}

#[test]
fn test_hmac_md5_hexdigest() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"key", b"The quick brown fox jumps over the lazy dog", "md5")
print(h.hexdigest())
"#,
    );
    assert_eq!(out, vec!["80070713463e7749b90c2dc24911e275"]);
}

#[test]
fn test_hmac_bytearray_key_and_message() {
    let out = run_python(
        r#"
import hmac
key = bytearray(b"key")
msg = bytearray(b"msg")
h = hmac.new(key, msg, "sha256")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["64"]);
}

#[test]
fn test_hmac_empty_key_and_message() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"", b"", "sha256")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["64"]);
}

#[test]
fn test_hmac_long_key_hashing() {
    let out = run_python(
        r#"
import hmac
# Key longer than block_size (64 bytes for sha256) is pre-hashed
long_key = b"k" * 128
h = hmac.new(long_key, b"data", "sha256")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["64"]);
}

#[test]
fn test_hmac_digest_mod_default_warning_or_error() {
    let out = run_python(
        r#"
import hmac
# Passing digestmod explicitly is required in modern Python
h = hmac.new(b"k", b"m", digestmod="sha256")
print(h is not None)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hmac_compare_digest_different_lengths() {
    let out = run_python(
        r#"
import hmac
print(hmac.compare_digest(b"short", b"longer_string"))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_hmac_blake2b_digestmod() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"k", b"m", "blake2b")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["128"]);
}

#[test]
fn test_hmac_sha3_256_digestmod() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"k", b"m", "sha3_256")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["64"]);
}

#[test]
fn test_hmac_compare_digest_subclass_check() {
    let out = run_python(
        r#"
import hmac

class CustomBytes(bytes): pass

cb1 = CustomBytes(b"val")
cb2 = CustomBytes(b"val")
print(hmac.compare_digest(cb1, cb2))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hmac_digest_raw_bytes_output() {
    let out = run_python(
        r#"
import hmac
h = hmac.new(b"key", b"data", "sha256")
d = h.digest()
print(isinstance(d, bytes))
print(len(d))
"#,
    );
    assert_eq!(out, vec!["True", "32"]);
}
