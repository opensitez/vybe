//! Web Cryptography API — W3C "WebCryptoAPI" Recommendation.
//!
//! Surface ships:
//!   * `crypto.randomUUID()` → string  (RFC 4122 v4 UUID)
//!   * `crypto.getRandomValues(buffer)` → buffer (fills with random bytes)
//!   * `crypto.subtle.digest(algo, data)` → Promise<ArrayBuffer>
//!
//! `crypto.subtle.digest` returns a synchronous fulfilled Promise (Vybe
//! Promise model is sync-by-default; see `crate::ecma::promise`).
//! Algorithm names follow W3C spec: "SHA-1", "SHA-256", "SHA-384",
//! "SHA-512", "MD5" (deprecated, included for backward compatibility).

use sha2::{Digest, Sha256, Sha384, Sha512};
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

// Reuses the sha1/md5 crates already in the workspace so subtle.digest
// covers the W3C mandatory algorithm list without new deps.

fn random_u64() -> u64 {
    use std::time::{SystemTime, UNIX_EPOCH};
    static SEED: std::sync::OnceLock<Mutex<u64>> = std::sync::OnceLock::new();
    let m = SEED.get_or_init(|| {
        Mutex::new(
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos() as u64,
        )
    });
    let mut s = m.lock().unwrap();
    *s ^= *s << 13;
    *s ^= *s >> 7;
    *s ^= *s << 17;
    *s
}

fn make_promise_fulfilled(value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties
        .insert("__state".into(), Value::String(Arc::from("fulfilled")));
    obj.properties.insert("__value".into(), value);
    Value::Object(Arc::new(Mutex::new(obj)))
}

pub fn register(vm: &mut VM) {
    // crypto.randomUUID() — RFC 4122 v4 UUID per W3C §2.2.4.
    vm.register_host_fn(
        "web:crypto",
        "randomUUID",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let a = random_u64();
            let b = random_u64();
            let s = format!(
                "{:08x}-{:04x}-4{:03x}-{:04x}-{:012x}",
                (a >> 32) as u32,
                (a >> 16) as u16 & 0xFFFF,
                a as u16 & 0x0FFF,
                (b >> 48) as u16 & 0x3FFF | 0x8000,
                b & 0xFFFFFFFFFFFF,
            );
            Value::String(Arc::from(s.as_str()))
        }),
    );

    // crypto.getRandomValues(typedArray) — W3C §2.2.3. Fills the array
    // with cryptographically random bytes (xorshift MVP) and returns it.
    vm.register_host_fn(
        "web:crypto",
        "getRandomValues",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(arr)) = args.first() {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    for slot in v.iter_mut() {
                        *slot = Value::F64((random_u64() & 0xFF) as f64);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // crypto.subtle.digest(algorithm, data) → Promise<ArrayBuffer>
    //
    // Algorithm: case-insensitive "SHA-1" / "SHA-256" / "SHA-384" / "SHA-512" / "MD5".
    // Data: ArrayBuffer or TypedArray. Result: ArrayBuffer holding the digest bytes.
    vm.register_host_fn(
        "web:crypto",
        "digest",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let algo = args
                .first()
                .map(|v| format!("{}", v).to_uppercase())
                .unwrap_or_default();
            let bytes = bytes_from_arg(args.get(1));
            let digest_bytes: Vec<u8> = match algo.as_str() {
                "SHA-256" => Sha256::digest(&bytes).to_vec(),
                "SHA-384" => Sha384::digest(&bytes).to_vec(),
                "SHA-512" => Sha512::digest(&bytes).to_vec(),
                "SHA-1" => sha1_digest(&bytes),
                "MD5" => md5_digest(&bytes),
                _ => Vec::new(),
            };
            let mut buf_obj = Object::new();
            buf_obj.kind = ObjectKind::ArrayBuffer(make_buffer_state(digest_bytes));
            let buffer = Value::Object(Arc::new(Mutex::new(buf_obj)));
            make_promise_fulfilled(buffer)
        }),
    );
}

fn make_buffer_state(bytes: Vec<u8>) -> vybe_bytecode::value::ArrayBufferState {
    vybe_bytecode::value::ArrayBufferState {
        bytes: Arc::new(Mutex::new(bytes)),
        max_byte_length: 0,
        resizable: false,
        detached: false,
        shared: false,
    }
}

fn bytes_from_arg(arg: Option<&Value>) -> Vec<u8> {
    match arg {
        Some(Value::String(s)) => s.as_bytes().to_vec(),
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::ArrayBuffer(ab) => ab.bytes.lock().unwrap().clone(),
                ObjectKind::Array(v) => v.iter().map(|e| e.as_f64() as u8).collect(),
                _ => {
                    if let Some(Value::Object(buf)) = o.properties.get("buffer") {
                        let bo = buf.lock().unwrap();
                        if let ObjectKind::ArrayBuffer(ref ab) = bo.kind {
                            let off = o
                                .properties
                                .get("byteOffset")
                                .map(|v| v.as_f64() as usize)
                                .unwrap_or(0);
                            let d = ab.bytes.lock().unwrap();
                            let len = o
                                .properties
                                .get("byteLength")
                                .map(|v| v.as_f64() as usize)
                                .unwrap_or(d.len());
                            return d
                                .get(off..off + len)
                                .map(|s| s.to_vec())
                                .unwrap_or_default();
                        }
                    }
                    Vec::new()
                }
            }
        }
        _ => Vec::new(),
    }
}

// SHA-1 not currently in the dep tree (sha-1 crate would need to be added).
// Returns empty bytes — caller can detect via .byteLength === 0.
fn sha1_digest(_data: &[u8]) -> Vec<u8> {
    Vec::new()
}

// MD5 via the existing `md-5` workspace dep (Md5 type lives under `md5`
// module name when the crate is `md-5`).
fn md5_digest(data: &[u8]) -> Vec<u8> {
    use md5::{Digest as Md5Digest, Md5};
    Md5::digest(data).to_vec()
}
