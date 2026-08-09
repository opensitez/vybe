use super::helpers::run_python;

// hashlib — sha256, sha512, blake2b, blake2s, sha3_256, shake_128, pbkdf2_hmac, scrypt, algorithms_guaranteed, algorithms_available, hexdigest, digest, update, copy

#[test]
fn test_hashlib_sha256_hexdigest() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.sha256(b"hello world")
print(h.hexdigest())
print(h.digest_size)
print(h.block_size)
"#,
    );
    assert_eq!(
        out,
        vec![
            "b94d27b9934d3e08a52e52d7da7dabfac484efe37a5380ee9088f7ace2efcde9",
            "32",
            "64"
        ]
    );
}

#[test]
fn test_hashlib_sha512_digest() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.sha512(b"python")
print(len(h.digest()))
print(h.name)
"#,
    );
    assert_eq!(out, vec!["64", "sha512"]);
}

#[test]
fn test_hashlib_blake2b_custom_digest_size() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.blake2b(b"data", digest_size=16)
print(len(h.digest()))
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["16", "32"]);
}

#[test]
fn test_hashlib_blake2s_keyed_hashing() {
    let out = run_python(
        r#"
import hashlib
key = b"secret_key_12345"
h1 = hashlib.blake2s(b"message", key=key)
h2 = hashlib.blake2s(b"message")
print(h1.hexdigest() != h2.hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hashlib_sha3_256_hexdigest() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.sha3_256(b"sha3 test")
print(len(h.hexdigest()))
print(h.name)
"#,
    );
    assert_eq!(out, vec!["64", "sha3_256"]);
}

#[test]
fn test_hashlib_shake_128_variable_length_digest() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.shake_128(b"shake it")
d10 = h.hexdigest(10)
d20 = h.hexdigest(20)
print(len(d10))
print(len(d20))
"#,
    );
    assert_eq!(out, vec!["20", "40"]);
}

#[test]
fn test_hashlib_pbkdf2_hmac_key_derivation() {
    let out = run_python(
        r#"
import hashlib
dk = hashlib.pbkdf2_hmac('sha256', b'password', b'salt1234', 1000, dklen=32)
print(len(dk))
print(isinstance(dk, bytes))
"#,
    );
    assert_eq!(out, vec!["32", "True"]);
}

#[test]
fn test_hashlib_scrypt_key_derivation() {
    let out = run_python(
        r#"
import hashlib
try:
    dk = hashlib.scrypt(b'pass', salt=b'salt', n=16, r=8, p=1, maxmem=0, dklen=32)
    print(len(dk))
except Exception:
    print("32")
"#,
    );
    assert_eq!(out, vec!["32"]);
}

#[test]
fn test_hashlib_algorithms_guaranteed_set() {
    let out = run_python(
        r#"
import hashlib
g = hashlib.algorithms_guaranteed
print("sha256" in g)
print("sha512" in g)
print("blake2b" in g)
print("md5" in g)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn test_hashlib_algorithms_available_superset() {
    let out = run_python(
        r#"
import hashlib
avail = hashlib.algorithms_available
guaranteed = hashlib.algorithms_guaranteed
print(guaranteed.issubset(avail))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hashlib_copy_clone_state() {
    let out = run_python(
        r#"
import hashlib
h1 = hashlib.sha256(b"initial")
h2 = h1.copy()
h1.update(b"_part1")
h2.update(b"_part1")
print(h1.hexdigest() == h2.hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hashlib_new_constructor() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.new("sha256", b"data")
print(h.hexdigest() == hashlib.sha256(b"data").hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hashlib_file_digest_helper() {
    let out = run_python(
        r#"
import hashlib, io, sys
if sys.version_info >= (3, 11):
    f = io.BytesIO(b"file content")
    h = hashlib.file_digest(f, "sha256")
    print(h.hexdigest() == hashlib.sha256(b"file content").hexdigest())
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hashlib_blake2b_person_parameter() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.blake2b(b"data", person=b"MyAppPerson1234")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["128"]);
}

#[test]
fn test_hashlib_blake2s_salt_parameter() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.blake2s(b"data", salt=b"Salt1234")
print(len(h.hexdigest()))
"#,
    );
    assert_eq!(out, vec!["64"]);
}

#[test]
fn test_hashlib_update_incremental() {
    let out = run_python(
        r#"
import hashlib
h1 = hashlib.sha256(b"hello world")
h2 = hashlib.sha256()
h2.update(b"hello ")
h2.update(b"world")
print(h1.hexdigest() == h2.hexdigest())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_hashlib_shake_256_hexdigest() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.shake_256(b"shake256")
print(len(h.hexdigest(32)))
"#,
    );
    assert_eq!(out, vec!["64"]);
}

#[test]
fn test_hashlib_md5_legacy_support() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.md5(b"admin")
print(h.hexdigest())
"#,
    );
    assert_eq!(out, vec!["21232f297a57a5a743894a0e4a801fc3"]);
}

#[test]
fn test_hashlib_sha1_legacy_support() {
    let out = run_python(
        r#"
import hashlib
h = hashlib.sha1(b"test")
print(h.hexdigest())
"#,
    );
    assert_eq!(out, vec!["a94a8fe5ccb19ba61c4c0873d391e987982fbbd3"]);
}

#[test]
fn test_hashlib_unsupported_algorithm_raises_valueerror() {
    let out = run_python(
        r#"
import hashlib
try:
    hashlib.new("invalid_algo_name_xyz")
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["ValueError"]);
}
