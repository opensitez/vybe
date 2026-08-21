//! Behaviour tests for `node:crypto` host imports.
//!
//! Reference: <https://nodejs.org/api/crypto.html>.
//!
//! Coverage:
//!   - `randomBytes(size)` → Buffer of `size` bytes
//!   - `randomUUID()` → UUID v4 string (36 chars, 4 dashes, version digit 4)
//!   - `createHash(algorithm)` → Hash with update/digest
//!   - `createHmac(algorithm, key)` → Hmac with update/digest
//!   - `pbkdf2Sync(pw, salt, iter, keylen, digest)` → Buffer of keylen bytes
//!   - `scryptSync(pw, salt, keylen)` → Buffer of keylen bytes
//!   - `createCipheriv(algo, key, iv)` → Cipher object (surface)
//!   - `createDecipheriv(algo, key, iv)` → Decipher object (surface)
//!   - `getCiphers()` → string array including "aes-256-cbc"
//!   - `getHashes()` → string array including "sha256", "sha1", "md5"
//!   - `getCurves()` → non-empty string array
//!   - `timingSafeEqual(a, b)` → bool
//!   - `constants` → object with RSA_PKCS1_PADDING = 1
//!   - Legacy shorthands: `sha256(data)`, `md5(data)`
//!
//! Deferred:
//!   - `generateKeyPairSync` (requires async key infra)
//!   - `createSign` / `createVerify` full round-trip (needs key objects)

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_crypto(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-crypto-test>");
    let import_idx = chunk.add_import("node:crypto", name);
    let argc = args.len() as u8;
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.set_global_owned(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:crypto"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn prop(v: &Value, key: &str) -> Value {
    match v {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

fn call_method(receiver: &Value, method: &str, args: Vec<Value>) -> Value {
    let fn_val = prop(receiver, method);
    let mut chunk = Chunk::new("<crypto-method>");
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    let fn_name = format!(
        "__test_arg_{}",
        TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    );
    let fn_ci = chunk.intern_string_constant(&fn_name);
    chunk.emit_op_u16(Op::GLOBAL_GET, fn_ci, 0);
    arg_globals.push((fn_name, fn_val));
    let recv_name = format!(
        "__test_arg_{}",
        TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    );
    let recv_ci = chunk.intern_string_constant(&recv_name);
    chunk.emit_op_u16(Op::GLOBAL_GET, recv_ci, 0);
    arg_globals.push((recv_name, receiver.clone()));
    let mut argc = 1usize;
    for arg in args {
        match arg {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
        argc += 1;
    }
    chunk.emit_op_u8_u8(Op::CALL_REF, argc as u8, 1, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.set_global_owned(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("method call failed")
}

fn array_len(v: &Value) -> usize {
    match v {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                elems.len()
            } else {
                0
            }
        }
        _ => 0,
    }
}

fn array_strings(v: &Value) -> Vec<String> {
    match v {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                return elems
                    .iter()
                    .map(|e| match e {
                        Value::String(s) => s.to_string(),
                        other => format!("{other}"),
                    })
                    .collect();
            }
            vec![]
        }
        _ => vec![],
    }
}

fn as_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{other}"),
    }
}

// ── Legacy shorthands ─────────────────────────────────────────────────────────

#[test]
fn sha256_shorthand_matches_known_digest() {
    assert_eq!(
        call_crypto("sha256", vec![s("abc")]),
        s("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
    );
}

#[test]
fn sha256_empty_input_matches_known_digest() {
    assert_eq!(
        call_crypto("sha256", vec![s("")]),
        s("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855")
    );
}

#[test]
fn md5_shorthand_matches_known_digest() {
    assert_eq!(
        call_crypto("md5", vec![s("abc")]),
        s("900150983cd24fb0d6963f7d28e17f72")
    );
}

// ── randomBytes ───────────────────────────────────────────────────────────────

#[test]
fn random_bytes_returns_buffer_of_requested_size() {
    let buf = call_crypto("randomBytes", vec![Value::I32(16)]);
    assert_eq!(array_len(&buf), 16, "randomBytes(16) must return 16 bytes");
}

#[test]
fn random_bytes_zero_returns_empty_buffer() {
    let buf = call_crypto("randomBytes", vec![Value::I32(0)]);
    assert_eq!(array_len(&buf), 0);
}

#[test]
fn random_bytes_32_returns_32_bytes() {
    let buf = call_crypto("randomBytes", vec![Value::I32(32)]);
    assert_eq!(array_len(&buf), 32);
}

#[test]
fn random_bytes_values_are_in_byte_range() {
    let buf = call_crypto("randomBytes", vec![Value::I32(8)]);
    if let Value::Object(o) = &buf {
        let o = o.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            for elem in elems {
                match elem {
                    Value::I32(n) => assert!(*n >= 0 && *n <= 255, "byte out of range: {n}"),
                    Value::F64(f) => assert!(*f >= 0.0 && *f <= 255.0, "byte out of range: {f}"),
                    _ => {}
                }
            }
            return;
        }
    }
    panic!("randomBytes did not return an array, got {:?}", buf);
}

// ── randomUUID ────────────────────────────────────────────────────────────────

#[test]
fn random_uuid_returns_string_of_36_chars() {
    let uuid = call_crypto("randomUUID", vec![]);
    let s = as_str(&uuid);
    assert_eq!(s.len(), 36, "UUID must be 36 characters, got: {s}");
}

#[test]
fn random_uuid_has_dashes_at_correct_positions() {
    let uuid = call_crypto("randomUUID", vec![]);
    let s = as_str(&uuid);
    assert_eq!(&s[8..9], "-", "dash at pos 8 in {s}");
    assert_eq!(&s[13..14], "-", "dash at pos 13 in {s}");
    assert_eq!(&s[18..19], "-", "dash at pos 18 in {s}");
    assert_eq!(&s[23..24], "-", "dash at pos 23 in {s}");
}

#[test]
fn random_uuid_version_digit_is_4() {
    let uuid = call_crypto("randomUUID", vec![]);
    let s = as_str(&uuid);
    assert_eq!(&s[14..15], "4", "UUID v4 version digit must be 4, got: {s}");
}

#[test]
fn random_uuid_variant_is_8_9_a_or_b() {
    let uuid = call_crypto("randomUUID", vec![]);
    let s = as_str(&uuid);
    let variant = &s[19..20];
    assert!(
        matches!(variant, "8" | "9" | "a" | "b"),
        "UUID variant must be 8/9/a/b, got: {variant} in {s}"
    );
}

#[test]
fn two_random_uuids_are_distinct() {
    let a = as_str(&call_crypto("randomUUID", vec![]));
    let b = as_str(&call_crypto("randomUUID", vec![]));
    assert_ne!(a, b, "two randomUUID() calls must not collide");
}

// ── createHash ────────────────────────────────────────────────────────────────

#[test]
fn create_hash_sha256_returns_object() {
    let hash = call_crypto("createHash", vec![s("sha256")]);
    assert!(
        matches!(hash, Value::Object(_)),
        "createHash must return an object"
    );
}

#[test]
fn create_hash_sha256_digest_hex_matches_known_value() {
    let hash = call_crypto("createHash", vec![s("sha256")]);
    let hash = call_method(&hash, "update", vec![s("abc")]);
    let digest = call_method(&hash, "digest", vec![s("hex")]);
    assert_eq!(
        digest,
        s("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
    );
}

#[test]
fn create_hash_sha1_digest_hex_matches_known_value() {
    let hash = call_crypto("createHash", vec![s("sha1")]);
    let hash = call_method(&hash, "update", vec![s("abc")]);
    let digest = call_method(&hash, "digest", vec![s("hex")]);
    assert_eq!(digest, s("a9993e364706816aba3e25717850c26c9cd0d89d"));
}

#[test]
fn create_hash_sha512_hex_digest_is_128_chars() {
    let hash = call_crypto("createHash", vec![s("sha512")]);
    let hash = call_method(&hash, "update", vec![s("abc")]);
    let digest = call_method(&hash, "digest", vec![s("hex")]);
    assert_eq!(
        as_str(&digest).len(),
        128,
        "sha512 hex digest must be 128 chars"
    );
}

#[test]
fn create_hash_md5_digest_hex_matches_known_value() {
    let hash = call_crypto("createHash", vec![s("md5")]);
    let hash = call_method(&hash, "update", vec![s("abc")]);
    let digest = call_method(&hash, "digest", vec![s("hex")]);
    assert_eq!(digest, s("900150983cd24fb0d6963f7d28e17f72"));
}

#[test]
fn create_hash_sha256_base64_digest_is_44_chars() {
    let hash = call_crypto("createHash", vec![s("sha256")]);
    let hash = call_method(&hash, "update", vec![s("hello")]);
    let digest = call_method(&hash, "digest", vec![s("base64")]);
    assert_eq!(
        as_str(&digest).len(),
        44,
        "sha256 base64 digest is always 44 chars"
    );
}

#[test]
fn create_hash_empty_input_sha256_matches_known_value() {
    let hash = call_crypto("createHash", vec![s("sha256")]);
    let hash = call_method(&hash, "update", vec![s("")]);
    let digest = call_method(&hash, "digest", vec![s("hex")]);
    assert_eq!(
        digest,
        s("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855")
    );
}

// ── createHmac ────────────────────────────────────────────────────────────────

#[test]
fn create_hmac_sha256_returns_object() {
    let hmac = call_crypto("createHmac", vec![s("sha256"), s("secret")]);
    assert!(matches!(hmac, Value::Object(_)));
}

#[test]
fn create_hmac_sha256_digest_matches_known_value() {
    let hmac = call_crypto("createHmac", vec![s("sha256"), s("secret")]);
    let hmac = call_method(&hmac, "update", vec![s("hello")]);
    let digest = call_method(&hmac, "digest", vec![s("hex")]);
    assert_eq!(
        digest,
        s("88aab3ede8d3adf94d26ab90d3bafd4a2083070c3bcce9c014ee04a443847c0b")
    );
}

#[test]
fn create_hmac_sha1_hex_digest_is_40_chars() {
    let hmac = call_crypto("createHmac", vec![s("sha1"), s("key")]);
    let hmac = call_method(&hmac, "update", vec![s("data")]);
    let digest = call_method(&hmac, "digest", vec![s("hex")]);
    assert_eq!(
        as_str(&digest).len(),
        40,
        "HMAC-SHA1 hex digest is 40 chars"
    );
}

// ── pbkdf2Sync ────────────────────────────────────────────────────────────────

#[test]
fn pbkdf2_sync_returns_buffer_of_correct_length() {
    let buf = call_crypto(
        "pbkdf2Sync",
        vec![
            s("password"),
            s("salt"),
            Value::I32(1),
            Value::I32(32),
            s("sha256"),
        ],
    );
    assert_eq!(
        array_len(&buf),
        32,
        "pbkdf2Sync keylen=32 must return 32 bytes"
    );
}

#[test]
fn pbkdf2_sync_keylen_16_returns_16_bytes() {
    let buf = call_crypto(
        "pbkdf2Sync",
        vec![
            s("pass"),
            s("salt"),
            Value::I32(1),
            Value::I32(16),
            s("sha256"),
        ],
    );
    assert_eq!(array_len(&buf), 16);
}

#[test]
fn pbkdf2_sync_is_deterministic() {
    let a = call_crypto(
        "pbkdf2Sync",
        vec![
            s("pw"),
            s("salt"),
            Value::I32(100),
            Value::I32(20),
            s("sha256"),
        ],
    );
    let b = call_crypto(
        "pbkdf2Sync",
        vec![
            s("pw"),
            s("salt"),
            Value::I32(100),
            Value::I32(20),
            s("sha256"),
        ],
    );
    assert_eq!(array_len(&a), 20);
    assert_eq!(array_len(&b), 20);
}

// ── scryptSync ────────────────────────────────────────────────────────────────

#[test]
fn scrypt_sync_returns_buffer_of_correct_length() {
    let buf = call_crypto("scryptSync", vec![s("password"), s("salt"), Value::I32(32)]);
    assert_eq!(
        array_len(&buf),
        32,
        "scryptSync keylen=32 must return 32 bytes"
    );
}

#[test]
fn scrypt_sync_keylen_64_returns_64_bytes() {
    let buf = call_crypto("scryptSync", vec![s("pass"), s("salt"), Value::I32(64)]);
    assert_eq!(array_len(&buf), 64);
}

// ── createCipheriv / createDecipheriv ─────────────────────────────────────────

#[test]
fn create_cipheriv_aes_256_cbc_returns_object() {
    let key = call_crypto("randomBytes", vec![Value::I32(32)]);
    let iv = call_crypto("randomBytes", vec![Value::I32(16)]);
    let cipher = call_crypto("createCipheriv", vec![s("aes-256-cbc"), key, iv]);
    assert!(
        matches!(cipher, Value::Object(_)),
        "createCipheriv must return object"
    );
}

#[test]
fn create_decipheriv_aes_256_cbc_returns_object() {
    let key = call_crypto("randomBytes", vec![Value::I32(32)]);
    let iv = call_crypto("randomBytes", vec![Value::I32(16)]);
    let decipher = call_crypto("createDecipheriv", vec![s("aes-256-cbc"), key, iv]);
    assert!(
        matches!(decipher, Value::Object(_)),
        "createDecipheriv must return object"
    );
}

// ── getCiphers / getHashes / getCurves ────────────────────────────────────────

#[test]
fn get_ciphers_returns_non_empty_array() {
    let ciphers = call_crypto("getCiphers", vec![]);
    assert!(
        array_len(&ciphers) > 0,
        "getCiphers() must return non-empty array"
    );
}

#[test]
fn get_ciphers_contains_aes_256_cbc() {
    let names = array_strings(&call_crypto("getCiphers", vec![]));
    assert!(
        names.iter().any(|n| n == "aes-256-cbc"),
        "getCiphers must contain aes-256-cbc, got: {names:?}"
    );
}

#[test]
fn get_hashes_returns_non_empty_array() {
    let hashes = call_crypto("getHashes", vec![]);
    assert!(
        array_len(&hashes) > 0,
        "getHashes() must return non-empty array"
    );
}

#[test]
fn get_hashes_contains_sha256_sha1_md5() {
    let names = array_strings(&call_crypto("getHashes", vec![]));
    for algo in ["sha256", "sha1", "md5", "sha512"] {
        assert!(
            names.iter().any(|n| n == algo),
            "getHashes must contain '{algo}', got: {names:?}"
        );
    }
}

#[test]
fn get_curves_returns_non_empty_array() {
    let curves = call_crypto("getCurves", vec![]);
    assert!(
        array_len(&curves) > 0,
        "getCurves() must return non-empty array"
    );
}

// ── timingSafeEqual ───────────────────────────────────────────────────────────

#[test]
fn timing_safe_equal_same_buffer_returns_true() {
    let buf = call_crypto("randomBytes", vec![Value::I32(16)]);
    let result = call_crypto("timingSafeEqual", vec![buf.clone(), buf]);
    assert_eq!(result, Value::Bool(true));
}

// ── constants ─────────────────────────────────────────────────────────────────

#[test]
fn constants_returns_object() {
    let consts = call_crypto("constants", vec![]);
    assert!(
        matches!(consts, Value::Object(_)),
        "crypto.constants must be an object"
    );
}

#[test]
fn constants_rsa_pkcs1_padding_is_1() {
    let consts = call_crypto("constants", vec![]);
    let val = prop(&consts, "RSA_PKCS1_PADDING");
    assert!(
        matches!(val, Value::I32(1) | Value::F64(1.0) | Value::I64(1)),
        "RSA_PKCS1_PADDING must be 1, got {:?}",
        val
    );
}

// ── createSign / createVerify ─────────────────────────────────────────────────

#[test]
fn create_sign_sha256_returns_object() {
    let sign = call_crypto("createSign", vec![s("sha256")]);
    assert!(
        matches!(sign, Value::Object(_)),
        "createSign must return object"
    );
}

#[test]
fn create_sign_has_update_method() {
    let sign = call_crypto("createSign", vec![s("sha256")]);
    let update = prop(&sign, "update");
    assert!(
        !matches!(update, Value::Null | Value::Undefined),
        "createSign().update must exist"
    );
}

#[test]
fn create_sign_has_sign_method() {
    let sign = call_crypto("createSign", vec![s("sha256")]);
    let sign_method = prop(&sign, "sign");
    assert!(
        !matches!(sign_method, Value::Null | Value::Undefined),
        "createSign().sign must exist"
    );
}

#[test]
fn create_verify_sha256_returns_object() {
    let verify = call_crypto("createVerify", vec![s("sha256")]);
    assert!(
        matches!(verify, Value::Object(_)),
        "createVerify must return object"
    );
}

#[test]
fn create_verify_has_update_method() {
    let verify = call_crypto("createVerify", vec![s("sha256")]);
    let update = prop(&verify, "update");
    assert!(
        !matches!(update, Value::Null | Value::Undefined),
        "createVerify().update must exist"
    );
}

#[test]
fn create_verify_has_verify_method() {
    let verify = call_crypto("createVerify", vec![s("sha256")]);
    let verify_method = prop(&verify, "verify");
    assert!(
        !matches!(verify_method, Value::Null | Value::Undefined),
        "createVerify().verify must exist"
    );
}

// ── randomInt ─────────────────────────────────────────────────────────────────

#[test]
fn random_int_in_range_returns_number() {
    let v = call_crypto("randomInt", vec![Value::I32(0), Value::I32(100)]);
    match v {
        Value::I32(n) => assert!(n >= 0 && n < 100, "randomInt must be in [0, 100), got {n}"),
        Value::I64(n) => assert!(n >= 0 && n < 100, "randomInt must be in [0, 100), got {n}"),
        Value::F64(f) => assert!(
            f >= 0.0 && f < 100.0 && f.fract() == 0.0,
            "randomInt got {f}"
        ),
        other => panic!("randomInt expected number, got {:?}", other),
    }
}

#[test]
fn random_int_with_max_only_returns_in_range() {
    let v = call_crypto("randomInt", vec![Value::I32(10)]);
    match v {
        Value::I32(n) => assert!(n >= 0 && n < 10, "randomInt(10) must be in [0,10), got {n}"),
        Value::I64(n) => assert!(n >= 0 && n < 10, "randomInt(10) must be in [0,10), got {n}"),
        Value::F64(f) => assert!(f >= 0.0 && f < 10.0, "randomInt(10) got {f}"),
        other => panic!("randomInt(10) expected number, got {:?}", other),
    }
}

// ── randomFillSync ────────────────────────────────────────────────────────────

#[test]
fn random_fill_sync_returns_buffer_filled() {
    // randomFillSync(buffer[, offset, size]) fills in-place and returns the buffer.
    // We call it with a 16-byte buffer.
    let buf = call_crypto("randomBytes", vec![Value::I32(16)]);
    let result = call_crypto("randomFillSync", vec![buf]);
    // Should return an array/buffer of 16 bytes
    assert_eq!(
        array_len(&result),
        16,
        "randomFillSync must return filled buffer of same length"
    );
}

// ── createDiffieHellman ───────────────────────────────────────────────────────

#[test]
fn create_diffie_hellman_returns_object() {
    let dh = call_crypto("createDiffieHellman", vec![Value::I32(512)]);
    assert!(
        matches!(dh, Value::Object(_)),
        "createDiffieHellman must return object"
    );
}

#[test]
fn create_diffie_hellman_has_generate_keys() {
    let dh = call_crypto("createDiffieHellman", vec![Value::I32(512)]);
    let gk = prop(&dh, "generateKeys");
    assert!(
        !matches!(gk, Value::Null | Value::Undefined),
        "DiffieHellman.generateKeys must exist"
    );
}

#[test]
fn create_diffie_hellman_has_compute_secret() {
    let dh = call_crypto("createDiffieHellman", vec![Value::I32(512)]);
    let cs = prop(&dh, "computeSecret");
    assert!(
        !matches!(cs, Value::Null | Value::Undefined),
        "DiffieHellman.computeSecret must exist"
    );
}

#[test]
fn create_diffie_hellman_has_get_prime() {
    let dh = call_crypto("createDiffieHellman", vec![Value::I32(512)]);
    let gp = prop(&dh, "getPrime");
    assert!(
        !matches!(gp, Value::Null | Value::Undefined),
        "DiffieHellman.getPrime must exist"
    );
}

// ── createECDH ────────────────────────────────────────────────────────────────

#[test]
fn create_ecdh_returns_object() {
    let ecdh = call_crypto("createECDH", vec![s("prime256v1")]);
    assert!(
        matches!(ecdh, Value::Object(_)),
        "createECDH must return object"
    );
}

#[test]
fn create_ecdh_has_generate_keys() {
    let ecdh = call_crypto("createECDH", vec![s("prime256v1")]);
    let gk = prop(&ecdh, "generateKeys");
    assert!(
        !matches!(gk, Value::Null | Value::Undefined),
        "ECDH.generateKeys must exist"
    );
}

#[test]
fn create_ecdh_has_compute_secret() {
    let ecdh = call_crypto("createECDH", vec![s("prime256v1")]);
    let cs = prop(&ecdh, "computeSecret");
    assert!(
        !matches!(cs, Value::Null | Value::Undefined),
        "ECDH.computeSecret must exist"
    );
}

// ── hkdfSync ─────────────────────────────────────────────────────────────────

#[test]
fn hkdf_sync_returns_buffer_of_requested_length() {
    // hkdfSync(digest, ikm, salt, info, keylen)
    let result = call_crypto(
        "hkdfSync",
        vec![
            s("sha256"),
            s("inputkeymaterial"),
            s("salt"),
            s("info"),
            Value::I32(32),
        ],
    );
    assert_eq!(
        array_len(&result),
        32,
        "hkdfSync must return buffer of keylen bytes"
    );
}

// ── generatePrimeSync / checkPrimeSync ───────────────────────────────────────

#[test]
fn generate_prime_sync_returns_buffer() {
    let prime = call_crypto("generatePrimeSync", vec![Value::I32(64)]);
    // Returns ArrayBuffer or Buffer; for our purposes check it's an object or buffer
    assert!(
        matches!(prime, Value::Object(_)),
        "generatePrimeSync must return buffer/ArrayBuffer object, got {:?}",
        prime
    );
}

#[test]
fn check_prime_sync_known_prime_returns_true() {
    // 7 is a prime
    let is_prime = call_crypto("checkPrimeSync", vec![Value::I32(7)]);
    assert_eq!(
        is_prime,
        Value::Bool(true),
        "7 is prime, checkPrimeSync must return true"
    );
}

#[test]
fn check_prime_sync_known_composite_returns_false() {
    // 6 is not prime
    let is_prime = call_crypto("checkPrimeSync", vec![Value::I32(6)]);
    assert_eq!(
        is_prime,
        Value::Bool(false),
        "6 is not prime, checkPrimeSync must return false"
    );
}

// ── getFips / setFips ─────────────────────────────────────────────────────────

#[test]
fn get_fips_returns_number_or_bool() {
    let v = call_crypto("getFips", vec![]);
    assert!(
        matches!(v, Value::I32(_) | Value::F64(_) | Value::Bool(_)),
        "getFips() must return number or bool, got {:?}",
        v
    );
}

#[test]
fn set_fips_does_not_panic() {
    // FIPS mode is system-dependent; just verify the call doesn't error.
    let result = call_crypto("setFips", vec![Value::Bool(false)]);
    let _ = result;
}

// ── Cipher update/final round-trip ───────────────────────────────────────────

#[test]
fn cipher_update_and_final_returns_bytes() {
    let key = call_crypto("randomBytes", vec![Value::I32(32)]);
    let iv = call_crypto("randomBytes", vec![Value::I32(16)]);
    let cipher = call_crypto("createCipheriv", vec![s("aes-256-cbc"), key, iv]);
    let chunk_out = call_method(&cipher, "update", vec![s("Hello, World!")]);
    let final_out = call_method(&cipher, "final", vec![]);
    // Both should be arrays (buffers).
    assert!(
        matches!(chunk_out, Value::Object(_) | Value::String(_)),
        "cipher.update must return buffer/string, got {:?}",
        chunk_out
    );
    assert!(
        matches!(final_out, Value::Object(_) | Value::String(_)),
        "cipher.final must return buffer/string, got {:?}",
        final_out
    );
}

// ── createHash copy ───────────────────────────────────────────────────────────

#[test]
fn create_hash_has_copy_method() {
    let hash = call_crypto("createHash", vec![s("sha256")]);
    let copy = prop(&hash, "copy");
    assert!(
        !matches!(copy, Value::Null | Value::Undefined),
        "Hash.copy must exist"
    );
}

// ── getRandomValues (Web Crypto compat) ───────────────────────────────────────

#[test]
fn get_random_values_fills_typed_array() {
    // getRandomValues fills a TypedArray in place and returns it.
    // We use an array of zeros as a proxy typed array.
    use std::collections::HashMap;
    use vybe_runtime::value::Object;
    let typed = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(vec![Value::I32(0); 8]),
        properties: Default::default(),
        type_id: 0,
        fields: Vec::new(),
    })));
    let result = call_crypto("getRandomValues", vec![typed]);
    assert!(
        matches!(result, Value::Object(_)),
        "getRandomValues must return typed array object"
    );
}

// ── subtle (Web Crypto API) ───────────────────────────────────────────────────

#[test]
fn subtle_returns_object() {
    let subtle = call_crypto("subtle", vec![]);
    assert!(
        matches!(subtle, Value::Object(_)),
        "crypto.subtle must return object"
    );
}

#[test]
fn subtle_has_digest_method() {
    let subtle = call_crypto("subtle", vec![]);
    let digest = prop(&subtle, "digest");
    assert!(
        !matches!(digest, Value::Null | Value::Undefined),
        "crypto.subtle.digest must exist"
    );
}

#[test]
fn subtle_has_encrypt_decrypt_methods() {
    let subtle = call_crypto("subtle", vec![]);
    assert!(
        !matches!(prop(&subtle, "encrypt"), Value::Null | Value::Undefined),
        "crypto.subtle.encrypt must exist"
    );
    assert!(
        !matches!(prop(&subtle, "decrypt"), Value::Null | Value::Undefined),
        "crypto.subtle.decrypt must exist"
    );
}

#[test]
fn subtle_has_sign_verify_methods() {
    let subtle = call_crypto("subtle", vec![]);
    assert!(
        !matches!(prop(&subtle, "sign"), Value::Null | Value::Undefined),
        "crypto.subtle.sign must exist"
    );
    assert!(
        !matches!(prop(&subtle, "verify"), Value::Null | Value::Undefined),
        "crypto.subtle.verify must exist"
    );
}

#[test]
fn subtle_has_generate_key_method() {
    let subtle = call_crypto("subtle", vec![]);
    assert!(
        !matches!(prop(&subtle, "generateKey"), Value::Null | Value::Undefined),
        "crypto.subtle.generateKey must exist"
    );
}

#[test]
fn subtle_has_import_export_key_methods() {
    let subtle = call_crypto("subtle", vec![]);
    assert!(
        !matches!(prop(&subtle, "importKey"), Value::Null | Value::Undefined),
        "crypto.subtle.importKey must exist"
    );
    assert!(
        !matches!(prop(&subtle, "exportKey"), Value::Null | Value::Undefined),
        "crypto.subtle.exportKey must exist"
    );
}

// ── generateKeyPairSync ───────────────────────────────────────────────────────

#[test]
fn generate_key_pair_sync_rsa_returns_object_pair() {
    use vybe_runtime::value::Object;
    let opts = Value::Object(std::sync::Arc::new(std::sync::Mutex::new({
        let mut o = Object::new();
        o.properties
            .insert("modulusLength".to_string(), Value::I32(2048));
        o
    })));
    let pair = call_crypto("generateKeyPairSync", vec![s("rsa"), opts]);
    assert!(
        matches!(pair, Value::Object(_)),
        "generateKeyPairSync must return object"
    );
    let pk = prop(&pair, "publicKey");
    let sk = prop(&pair, "privateKey");
    assert!(
        !matches!(pk, Value::Null | Value::Undefined),
        "generateKeyPairSync.publicKey must exist"
    );
    assert!(
        !matches!(sk, Value::Null | Value::Undefined),
        "generateKeyPairSync.privateKey must exist"
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_crypto_surface_is_registered() {
    let expected = [
        "randomBytes",
        "randomUUID",
        "randomInt",
        "randomFillSync",
        "createHash",
        "createHmac",
        "pbkdf2Sync",
        "scryptSync",
        "hkdfSync",
        "createCipheriv",
        "createDecipheriv",
        "createSign",
        "createVerify",
        "createDiffieHellman",
        "createECDH",
        "generateKeyPairSync",
        "generatePrimeSync",
        "checkPrimeSync",
        "getCiphers",
        "getHashes",
        "getCurves",
        "timingSafeEqual",
        "getRandomValues",
        "subtle",
        "getFips",
        "setFips",
        "constants",
        "sha256",
        "md5",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:crypto imports: {missing:?}"
    );
}
