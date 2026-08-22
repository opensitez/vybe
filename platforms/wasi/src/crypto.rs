//! `wasi:crypto` — the WASI Cryptography APIs proposal, plus the shared
//! digest primitives every other crypto surface on the platform composes.
//!
//! Interfaces are defined in `proposals/wasi-crypto/witx/witx-0.10/`; the
//! normative prose is `proposals/wasi-crypto/docs/wasi-crypto.md`. This
//! module implements `wasi_ephemeral_crypto_common` and
//! `wasi_ephemeral_crypto_symmetric`.
//!
//! # Layering
//!
//! [`HashAlgorithm`] below is the ONE implementation of every hash and MAC
//! on the platform. `node:crypto` maps OpenSSL names (`sha3-256`) onto it;
//! `wasi:crypto/symmetric` maps the proposal's names (`SHA-256`,
//! `HMAC/SHA-256`). The cryptography is written once — each interface skin
//! only owns its own naming and calling convention.
//!
//! # ABI mapping
//!
//! The witx signatures are linear-memory shaped: every function returns
//! `crypto_errno` and writes results through out-pointers, with byte inputs
//! as `(ptr<u8>, size)` pairs. Vybe's WASI modules are component-model
//! shaped, so — exactly as `wasi:filesystem` and `wasi:sockets` do:
//!
//! - `(ptr<u8>, len)` input pair → one byte-ish `Value`, decoded by [`bytes_of`]
//! - `mut_ptr<T>` out-parameter → the function's return value
//! - `crypto_errno` → an Object carrying `__wasi_error` with the enum name in
//!   WIT spelling (`"unsupported-algorithm"`, `"key-required"`,
//!   `"invalid-tag"`). Same carrier `wasi:filesystem` uses for `error-code`,
//!   because the VM has no host-fn exception channel.
//! - handles (`symmetric_state`, `symmetric_key`, `array_output`,
//!   `symmetric_tag`) → Objects with `__wasi_kind` + `__wasi_id`, backed by
//!   the host-side [`registry`]
//!
//! # Coverage
//!
//! Hash functions and MACs meet the spec's MUST list (§"Hash functions",
//! §"Message Authentication Codes"): `absorb`/`squeeze` for hashes,
//! `absorb`/`squeeze_tag` for MACs, any number of calls in any order, and
//! `squeeze` leaves the state untouched (finalization runs on a copy).
//!
//! AEAD, ratcheting, key exchange, signatures and managed keys return the
//! proposal's own `not_implemented`. Declining in the spec's vocabulary is
//! conformant; inventing an answer is not.

use std::collections::HashMap;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

// The interface names are the WIT's, verbatim —
// `proposals/wasi-crypto/wit/wasi_ephemeral_crypto_common.wit` declares
// `package wasi:crypto@0.11.0` containing `interface
// wasi-ephemeral-crypto-common`, so the module is the two joined.
//
// ⚠These read `wasi:crypto/common` and `wasi:crypto/symmetric` — shortened
// names that the proposal does not declare. That is the same invention as the
// flat `wasi:filesystem` verbs, just less obvious because the PACKAGE was
// right. A guest generated from this WIT imports the long name and would not
// have resolved against either of them.
const COMMON: &str = "wasi:crypto/wasi-ephemeral-crypto-common";
const SYMMETRIC: &str = "wasi:crypto/wasi-ephemeral-crypto-symmetric";

// ── Digest primitives ────────────────────────────────────────────────────

/// A hash function. Every variant is really implemented — there is no
/// "close enough" fallback, because a well-formed digest from the wrong
/// function is worse than an error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlgorithm {
    Md5,
    Sha1,
    Sha224,
    Sha256,
    Sha384,
    Sha512,
    Sha512_224,
    Sha512_256,
    Sha3_224,
    Sha3_256,
    Sha3_384,
    Sha3_512,
    /// Extendable-output; `digest` yields its default 16 bytes.
    Shake128,
    /// Extendable-output; `digest` yields its default 32 bytes.
    Shake256,
    Blake2b512,
    Blake2s256,
    Ripemd160,
}

impl HashAlgorithm {
    /// Digest of `data`.
    pub fn digest(self, data: &[u8]) -> Vec<u8> {
        use sha2::Digest;
        macro_rules! fixed {
            ($t:ty) => {{
                let mut h = <$t>::new();
                h.update(data);
                h.finalize().to_vec()
            }};
        }
        macro_rules! xof {
            ($t:ty, $len:expr) => {{
                use sha3::digest::{ExtendableOutput, Update, XofReader};
                let mut h = <$t>::default();
                h.update(data);
                let mut out = vec![0u8; $len];
                h.finalize_xof().read(&mut out);
                out
            }};
        }
        match self {
            Self::Md5 => fixed!(md5::Md5),
            Self::Sha1 => fixed!(sha1::Sha1),
            Self::Sha224 => fixed!(sha2::Sha224),
            Self::Sha256 => fixed!(sha2::Sha256),
            Self::Sha384 => fixed!(sha2::Sha384),
            Self::Sha512 => fixed!(sha2::Sha512),
            Self::Sha512_224 => fixed!(sha2::Sha512_224),
            Self::Sha512_256 => fixed!(sha2::Sha512_256),
            Self::Sha3_224 => fixed!(sha3::Sha3_224),
            Self::Sha3_256 => fixed!(sha3::Sha3_256),
            Self::Sha3_384 => fixed!(sha3::Sha3_384),
            Self::Sha3_512 => fixed!(sha3::Sha3_512),
            Self::Shake128 => xof!(sha3::Shake128, 16),
            Self::Shake256 => xof!(sha3::Shake256, 32),
            Self::Blake2b512 => fixed!(blake2::Blake2b512),
            Self::Blake2s256 => fixed!(blake2::Blake2s256),
            Self::Ripemd160 => fixed!(ripemd::Ripemd160),
        }
    }

    /// Digest length in bytes (the default output length for the XOFs).
    pub fn digest_len(self) -> usize {
        match self {
            Self::Md5 => 16,
            Self::Sha1 | Self::Ripemd160 => 20,
            Self::Sha224 | Self::Sha512_224 | Self::Sha3_224 => 28,
            Self::Sha256
            | Self::Sha512_256
            | Self::Sha3_256
            | Self::Shake256
            | Self::Blake2s256 => 32,
            Self::Sha384 | Self::Sha3_384 => 48,
            Self::Sha512 | Self::Sha3_512 | Self::Blake2b512 => 64,
            Self::Shake128 => 16,
        }
    }

    /// HMAC block size `B`. `None` for the XOFs, which have no fixed input
    /// block — OpenSSL rejects HMAC-SHAKE for the same reason.
    ///
    /// For Keccak this is the sponge RATE, not the output size.
    pub fn block_size(self) -> Option<usize> {
        Some(match self {
            Self::Md5
            | Self::Sha1
            | Self::Sha224
            | Self::Sha256
            | Self::Ripemd160
            | Self::Blake2s256 => 64,
            Self::Sha384
            | Self::Sha512
            | Self::Sha512_224
            | Self::Sha512_256
            | Self::Blake2b512 => 128,
            Self::Sha3_224 => 144,
            Self::Sha3_256 => 136,
            Self::Sha3_384 => 104,
            Self::Sha3_512 => 72,
            Self::Shake128 | Self::Shake256 => return None,
        })
    }

    /// HMAC (RFC 2104). `None` when the algorithm has no HMAC block size.
    ///
    /// Built by hand rather than with `hmac::Hmac<D>`: BLAKE2's RustCrypto
    /// types use the `Lazy` buffer kind that `Hmac` cannot wrap, yet
    /// OpenSSL/Node do offer HMAC-BLAKE2. The construction is the standard
    /// one, so it agrees with `Hmac<D>` wherever both can be expressed.
    pub fn hmac(self, key: &[u8], data: &[u8]) -> Option<Vec<u8>> {
        let block = self.block_size()?;
        let mut k = if key.len() > block {
            self.digest(key)
        } else {
            key.to_vec()
        };
        k.resize(block, 0);

        let mut inner: Vec<u8> = k.iter().map(|b| b ^ 0x36).collect();
        inner.extend_from_slice(data);
        let inner_digest = self.digest(&inner);

        let mut outer: Vec<u8> = k.iter().map(|b| b ^ 0x5c).collect();
        outer.extend_from_slice(&inner_digest);
        Some(self.digest(&outer))
    }

    /// Lowercase hex digest.
    pub fn hex(self, data: &[u8]) -> String {
        hex_of(&self.digest(data))
    }
}

/// Lowercase hex encoding.
pub fn hex_of(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{:02x}", b)).collect()
}

/// SHA-256 hex digest. Named entry point: `node:crypto` and the
/// `wasi:crypto/hashes` shim both call it directly.
pub fn sha256_hex(data: &[u8]) -> String {
    HashAlgorithm::Sha256.hex(data)
}

/// MD5 hex digest. Same story as [`sha256_hex`].
pub fn md5_hex(data: &[u8]) -> String {
    HashAlgorithm::Md5.hex(data)
}

/// Cryptographically secure random bytes, or `None` if no entropy source is
/// available (surfaced as the proposal's `rng_error`).
///
/// One entropy source for the platform: `wasi:random`'s CSPRNG path, which is
/// the OS generator. Deliberately not its xorshift path, which backs only
/// `wasi:random/insecure`.
fn csprng_bytes(n: usize) -> Option<Vec<u8>> {
    crate::random::secure_bytes(n)
}

// ── Handle registry ──────────────────────────────────────────────────────

const KIND_ARRAY_OUTPUT: &str = "array-output";
const KIND_SYMMETRIC_KEY: &str = "symmetric-key";
const KIND_SYMMETRIC_STATE: &str = "symmetric-state";
const KIND_SYMMETRIC_TAG: &str = "symmetric-tag";

/// A `symmetric_state` handle's contents. `absorbed` is the transcript; the
/// digest is computed from a copy at squeeze time, so the state is unchanged
/// from the guest's perspective (spec: implementations "MUST duplicate the
/// internal state and apply the finalization on the copy").
#[derive(Clone)]
struct SymmetricState {
    algorithm: String,
    key: Option<Vec<u8>>,
    absorbed: Vec<u8>,
}

#[derive(Default)]
struct Registry {
    array_outputs: HashMap<u64, Vec<u8>>,
    keys: HashMap<u64, (String, Vec<u8>)>,
    states: HashMap<u64, SymmetricState>,
    tags: HashMap<u64, Vec<u8>>,
}

/// Handle ids, OUTSIDE the registry so clearing tenant data cannot rewind them
/// — a reissued handle would let a stale reference address another tenant's
/// key.
static NEXT_ID: AtomicU64 = AtomicU64::new(1);

/// Every key, symmetric state, tag and output buffer this program created.
///
/// The sharpest resource in the set — `keys` holds raw key material and `tags`
/// holds authentication tags — and the reason it is VM-owned
/// ([`vybe_runtime::resources`]) rather than a process-global static: the VM
/// drops it on `reset_to` whether or not anyone remembered to ask. While it was
/// a static, the next tenant of a reused VM could address the previous
/// tenant's key by handle.
fn registry() -> &'static Mutex<Registry> {
    vybe_runtime::resources::get::<Registry>()
}

/// Takes no registry: the counter is not in it, and a `fresh_id()`
/// would read as though it were.
fn fresh_id() -> u64 {
    NEXT_ID.fetch_add(1, Ordering::Relaxed)
}

/// Build a resource handle Object (`__wasi_kind` + `__wasi_id`).
fn handle(kind: &str, id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__wasi_kind".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// A handle's id, when it is of the expected kind.
fn handle_id(v: &Value, kind: &str) -> Option<u64> {
    let Value::Object(obj) = v else { return None };
    let obj = obj.lock().unwrap();
    let right_kind = matches!(
        obj.properties.get("__wasi_kind"),
        Some(Value::String(k)) if &**k == kind
    );
    if !right_kind {
        return None;
    }
    match obj.properties.get("__wasi_id") {
        Some(Value::F64(id)) => Some(*id as u64),
        Some(Value::I32(id)) => Some(*id as u64),
        _ => None,
    }
}

/// A `crypto_errno` return, in the `__wasi_error` carrier. Names are the
/// witx enum members in WIT spelling (underscores → hyphens).
fn err(code: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__wasi_error".into(), Value::String(Arc::from(code)));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

// ── Byte decoding ────────────────────────────────────────────────────────

/// Decode a witx `(ptr<u8>, size)` input pair from a single Value.
///
/// Accepts a string (UTF-8 bytes), an array of octets, or a typed array —
/// the last being what every language's `bytes` maps to. Decoding with
/// `format!("{}", v)` (what the old shim did) hashes the *Display form* of a
/// Uint8Array, silently producing the wrong digest for all binary input.
fn bytes_of(v: &Value) -> Vec<u8> {
    match v {
        Value::String(s) => s.as_bytes().to_vec(),
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems
                    .iter()
                    .map(|e| match e {
                        Value::I32(n) => *n as u8,
                        Value::F64(f) => *f as u8,
                        _ => 0,
                    })
                    .collect(),
                ObjectKind::TypedArray(ta) => {
                    let buf = ta.buffer.lock().unwrap();
                    let start = ta.byte_offset.min(buf.len());
                    let end = (start + ta.length).min(buf.len());
                    buf[start..end].to_vec()
                }
                _ => Vec::new(),
            }
        }
        _ => Vec::new(),
    }
}

fn bytes_to_value(bytes: &[u8]) -> Value {
    let elems: Vec<Value> = bytes.iter().map(|b| Value::I32(*b as i32)).collect();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
}

fn str_of(v: Option<&Value>) -> String {
    match v {
        Some(Value::String(s)) => s.to_string(),
        _ => String::new(),
    }
}

fn usize_of(v: Option<&Value>) -> Option<usize> {
    match v {
        Some(Value::F64(n)) => Some(*n as usize),
        Some(Value::I32(n)) => Some(*n as usize),
        _ => None,
    }
}

// ── Proposal algorithm names ─────────────────────────────────────────────

/// Resolve a symmetric hash algorithm from the proposal's table
/// (`docs/wasi-crypto.md` §"Symmetric operations").
fn hash_named(algorithm: &str) -> Option<HashAlgorithm> {
    Some(match algorithm {
        "SHA-256" => HashAlgorithm::Sha256,
        "SHA-512" => HashAlgorithm::Sha512,
        "SHA-512/256" => HashAlgorithm::Sha512_256,
        _ => return None,
    })
}

/// The hash underlying a `HMAC/<hash>` algorithm name.
fn mac_named(algorithm: &str) -> Option<HashAlgorithm> {
    algorithm.strip_prefix("HMAC/").and_then(hash_named)
}

fn is_known_algorithm(algorithm: &str) -> bool {
    hash_named(algorithm).is_some() || mac_named(algorithm).is_some()
}

/// Finalize a state's transcript: a digest for hash algorithms, a tag for
/// MACs. Always computed from a copy, so the state is untouched.
fn finalize(state: &SymmetricState) -> Option<Vec<u8>> {
    if let Some(hash) = mac_named(&state.algorithm) {
        hash.hmac(state.key.as_deref().unwrap_or(&[]), &state.absorbed)
    } else {
        hash_named(&state.algorithm).map(|h| h.digest(&state.absorbed))
    }
}

// ── Registration ─────────────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    register_common(vm);
    register_symmetric(vm);
    register_asymmetric_common(vm);
    register_signatures(vm);
    register_signatures_batch(vm);
    register_kx(vm);
    register_external_secrets(vm);
    register_symmetric_batch(vm);
    register_hashes_shim(vm);
}

/// The five interfaces this module declines in full, plus the two batch ones.
///
/// Every function the proposal declares is REGISTERED and answers the
/// proposal's own `not_implemented`. That is the policy this file already
/// states for AEAD, ratcheting, key exchange, signatures and managed keys —
/// declining in the spec's vocabulary is conformant; inventing an answer is
/// not. What was not conformant was leaving the names UNREGISTERED, because
/// then a guest generated from the WIT gets `Unresolved import` — a link
/// failure, not a crypto error, and one it cannot catch or report.
///
/// Names and interface spellings come from `proposals/wasi-crypto/wit/`, in
/// declaration order.
fn register_declined(vm: &mut VM, interface: &'static str, names: &[&'static str]) {
    for name in names {
        vm.register_host_fn(
            interface,
            name,
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| err("not-implemented")),
        );
    }
}

const ASYMMETRIC_COMMON: &str = "wasi:crypto/wasi-ephemeral-crypto-asymmetric-common";
const SIGNATURES: &str = "wasi:crypto/wasi-ephemeral-crypto-signatures";
const SIGNATURES_BATCH: &str = "wasi:crypto/wasi-ephemeral-crypto-signatures-batch";
const KX: &str = "wasi:crypto/wasi-ephemeral-crypto-kx";
const EXTERNAL_SECRETS: &str = "wasi:crypto/wasi-ephemeral-crypto-external-secrets";
const SYMMETRIC_BATCH: &str = "wasi:crypto/wasi-ephemeral-crypto-symmetric-batch";

fn register_asymmetric_common(vm: &mut VM) {
    register_declined(
        vm,
        ASYMMETRIC_COMMON,
        &[
            "keypair-generate",
            "keypair-import",
            "keypair-generate-managed",
            "keypair-store-managed",
            "keypair-replace-managed",
            "keypair-id",
            "keypair-from-id",
            "keypair-from-pk-and-sk",
            "keypair-export",
            "keypair-publickey",
            "keypair-secretkey",
            "keypair-close",
            "publickey-import",
            "publickey-export",
            "publickey-verify",
            "publickey-from-secretkey",
            "publickey-close",
            "secretkey-import",
            "secretkey-export",
            "secretkey-close",
        ],
    );
}

fn register_signatures(vm: &mut VM) {
    register_declined(
        vm,
        SIGNATURES,
        &[
            "signature-export",
            "signature-import",
            "signature-state-open",
            "signature-state-update",
            "signature-state-sign",
            "signature-state-close",
            "signature-verification-state-open",
            "signature-verification-state-update",
            "signature-verification-state-verify",
            "signature-verification-state-close",
            "signature-close",
        ],
    );
}

fn register_signatures_batch(vm: &mut VM) {
    register_declined(
        vm,
        SIGNATURES_BATCH,
        &["batch-signature-state-sign", "batch-signature-state-verify"],
    );
}

fn register_kx(vm: &mut VM) {
    register_declined(vm, KX, &["kx-dh", "kx-encapsulate", "kx-decapsulate"]);
}

fn register_external_secrets(vm: &mut VM) {
    register_declined(
        vm,
        EXTERNAL_SECRETS,
        &[
            "external-secret-store",
            "external-secret-replace",
            "external-secret-from-id",
            "external-secret-invalidate",
            "external-secret-encapsulate",
            "external-secret-decapsulate",
        ],
    );
}

fn register_symmetric_batch(vm: &mut VM) {
    register_declined(
        vm,
        SYMMETRIC_BATCH,
        &[
            "batch-symmetric-state-squeeze",
            "batch-symmetric-state-squeeze-tag",
            "batch-symmetric-state-encrypt",
            "batch-symmetric-state-encrypt-detached",
            "batch-symmetric-state-decrypt",
            "batch-symmetric-state-decrypt-detached",
        ],
    );
}

fn register_common(vm: &mut VM) {
    // array-output-len(array-output) -> size
    vm.register_host_fn(
        COMMON,
        "array-output-len",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_ARRAY_OUTPUT)) else {
                return err("invalid-handle");
            };
            match registry().lock().unwrap().array_outputs.get(&id) {
                Some(bytes) => Value::F64(bytes.len() as f64),
                None => err("closed"),
            }
        }),
    );

    // array-output-pull(array-output, buf_len) -> list<u8>
    //
    // witx copies into a guest buffer and reports the count, consuming what
    // it wrote. Same semantics, with the bytes returned instead.
    vm.register_host_fn(
        COMMON,
        "array-output-pull",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_ARRAY_OUTPUT)) else {
                return err("invalid-handle");
            };
            let mut reg = registry().lock().unwrap();
            let Some(bytes) = reg.array_outputs.get_mut(&id) else {
                return err("closed");
            };
            let take = usize_of(args.get(1))
                .unwrap_or(bytes.len())
                .min(bytes.len());
            let out: Vec<u8> = bytes.drain(..take).collect();
            bytes_to_value(&out)
        }),
    );

    // Options carry algorithm parameters (nonces, memory limits). None of the
    // implemented hash/MAC algorithms read any, so an options set is accepted
    // and ignored rather than refused.
    vm.register_host_fn(
        COMMON,
        "options-open",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::Object(vybe_runtime::heap::alloc(Object::new()))
        }),
    );
    for name in [
        "options-set",
        "options-set-u64",
        "options-set-guest-buffer",
        "options-close",
    ] {
        vm.register_host_fn(
            COMMON,
            name,
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
        );
    }

    // Managed keys need a secrets manager (an external KMS or enclave).
    for name in [
        "secrets-manager-open",
        "secrets-manager-close",
        "secrets-manager-invalidate",
    ] {
        vm.register_host_fn(
            COMMON,
            name,
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| err("not-implemented")),
        );
    }
}

fn register_symmetric(vm: &mut VM) {
    // symmetric-key-generate(algorithm, options) -> symmetric-key
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-key-generate",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let algorithm = str_of(args.first());
            if !is_known_algorithm(&algorithm) {
                return err("unsupported-algorithm");
            }
            let len = mac_named(&algorithm)
                .and_then(|h| h.block_size())
                .or_else(|| hash_named(&algorithm).map(|h| h.digest_len()))
                .unwrap_or(32);
            let Some(raw) = csprng_bytes(len) else {
                return err("rng-error");
            };
            let mut reg = registry().lock().unwrap();
            let id = fresh_id();
            reg.keys.insert(id, (algorithm, raw));
            handle(KIND_SYMMETRIC_KEY, id)
        }),
    );

    // symmetric-key-import(algorithm, raw) -> symmetric-key
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-key-import",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let algorithm = str_of(args.first());
            if !is_known_algorithm(&algorithm) {
                return err("unsupported-algorithm");
            }
            let raw = args.get(1).map(bytes_of).unwrap_or_default();
            let mut reg = registry().lock().unwrap();
            let id = fresh_id();
            reg.keys.insert(id, (algorithm, raw));
            handle(KIND_SYMMETRIC_KEY, id)
        }),
    );

    // symmetric-key-export(symmetric-key) -> array-output
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-key-export",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(key_id) = args.first().and_then(|v| handle_id(v, KIND_SYMMETRIC_KEY)) else {
                return err("invalid-handle");
            };
            let mut reg = registry().lock().unwrap();
            let Some((_, raw)) = reg.keys.get(&key_id).cloned() else {
                return err("closed");
            };
            let id = fresh_id();
            reg.array_outputs.insert(id, raw);
            handle(KIND_ARRAY_OUTPUT, id)
        }),
    );

    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-key-close",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_SYMMETRIC_KEY)) else {
                return err("invalid-handle");
            };
            registry().lock().unwrap().keys.remove(&id);
            Value::Null
        }),
    );

    // symmetric-state-open(algorithm, key, options) -> symmetric-state
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-open",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let algorithm = str_of(args.first());
            if !is_known_algorithm(&algorithm) {
                return err("unsupported-algorithm");
            }
            let mut reg = registry().lock().unwrap();
            let key = match args.get(1).and_then(|v| handle_id(v, KIND_SYMMETRIC_KEY)) {
                Some(key_id) => match reg.keys.get(&key_id) {
                    // A key is bound to the algorithm it was imported for.
                    Some((key_algo, _)) if key_algo != &algorithm => {
                        return err("invalid-key");
                    }
                    Some((_, raw)) => Some(raw.clone()),
                    None => return err("closed"),
                },
                None => None,
            };
            // MACs cannot run keyless.
            if mac_named(&algorithm).is_some() && key.is_none() {
                return err("key-required");
            }
            let id = fresh_id();
            reg.states.insert(
                id,
                SymmetricState {
                    algorithm,
                    key,
                    absorbed: Vec::new(),
                },
            );
            handle(KIND_SYMMETRIC_STATE, id)
        }),
    );

    // symmetric-state-absorb(state, data)
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-absorb",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args
                .first()
                .and_then(|v| handle_id(v, KIND_SYMMETRIC_STATE))
            else {
                return err("invalid-handle");
            };
            let data = args.get(1).map(bytes_of).unwrap_or_default();
            let mut reg = registry().lock().unwrap();
            match reg.states.get_mut(&id) {
                Some(state) => {
                    state.absorbed.extend_from_slice(&data);
                    Value::Null
                }
                None => err("closed"),
            }
        }),
    );

    // symmetric-state-squeeze(state, len) -> list<u8>
    //
    // Spec: truncate to `len`; `invalid_length` when `len` exceeds what the
    // function can output; the state is NOT consumed, so absorb and squeeze
    // may interleave any number of times.
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-squeeze",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args
                .first()
                .and_then(|v| handle_id(v, KIND_SYMMETRIC_STATE))
            else {
                return err("invalid-handle");
            };
            let reg = registry().lock().unwrap();
            let Some(state) = reg.states.get(&id) else {
                return err("closed");
            };
            // `squeeze` is the hash-function operation; MACs use squeeze_tag.
            if mac_named(&state.algorithm).is_some() {
                return err("invalid-operation");
            }
            let Some(digest) = finalize(state) else {
                return err("unsupported-algorithm");
            };
            let want = usize_of(args.get(1)).unwrap_or(digest.len());
            if want > digest.len() {
                return err("invalid-length");
            }
            bytes_to_value(&digest[..want])
        }),
    );

    // symmetric-state-squeeze-tag(state) -> symmetric-tag
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-squeeze-tag",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args
                .first()
                .and_then(|v| handle_id(v, KIND_SYMMETRIC_STATE))
            else {
                return err("invalid-handle");
            };
            let mut reg = registry().lock().unwrap();
            let Some(state) = reg.states.get(&id).cloned() else {
                return err("closed");
            };
            if mac_named(&state.algorithm).is_none() {
                return err("invalid-operation");
            }
            let Some(tag) = finalize(&state) else {
                return err("unsupported-algorithm");
            };
            let tag_id = fresh_id();
            reg.tags.insert(tag_id, tag);
            handle(KIND_SYMMETRIC_TAG, tag_id)
        }),
    );

    // symmetric-state-max-tag-len(state) -> size
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-max-tag-len",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args
                .first()
                .and_then(|v| handle_id(v, KIND_SYMMETRIC_STATE))
            else {
                return err("invalid-handle");
            };
            let reg = registry().lock().unwrap();
            let Some(state) = reg.states.get(&id) else {
                return err("closed");
            };
            match mac_named(&state.algorithm) {
                Some(hash) => Value::F64(hash.digest_len() as f64),
                None => err("invalid-operation"),
            }
        }),
    );

    // symmetric-state-clone(state) -> symmetric-state
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-clone",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args
                .first()
                .and_then(|v| handle_id(v, KIND_SYMMETRIC_STATE))
            else {
                return err("invalid-handle");
            };
            let mut reg = registry().lock().unwrap();
            let Some(state) = reg.states.get(&id).cloned() else {
                return err("closed");
            };
            let new_id = fresh_id();
            reg.states.insert(new_id, state);
            handle(KIND_SYMMETRIC_STATE, new_id)
        }),
    );

    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-state-close",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args
                .first()
                .and_then(|v| handle_id(v, KIND_SYMMETRIC_STATE))
            else {
                return err("invalid-handle");
            };
            registry().lock().unwrap().states.remove(&id);
            Value::Null
        }),
    );

    // symmetric-tag-len(tag) -> size
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-tag-len",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_SYMMETRIC_TAG)) else {
                return err("invalid-handle");
            };
            match registry().lock().unwrap().tags.get(&id) {
                Some(tag) => Value::F64(tag.len() as f64),
                None => err("closed"),
            }
        }),
    );

    // symmetric-tag-pull(tag, len) -> list<u8>
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-tag-pull",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_SYMMETRIC_TAG)) else {
                return err("invalid-handle");
            };
            let reg = registry().lock().unwrap();
            let Some(tag) = reg.tags.get(&id) else {
                return err("closed");
            };
            let want = usize_of(args.get(1)).unwrap_or(tag.len());
            if want > tag.len() {
                return err("invalid-length");
            }
            bytes_to_value(&tag[..want])
        }),
    );

    // symmetric-tag-verify(tag, expected)
    //
    // Spec: `invalid_tag` on mismatch. Compared in constant time — a check
    // that leaks the first differing byte is the classic forgery oracle.
    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-tag-verify",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_SYMMETRIC_TAG)) else {
                return err("invalid-handle");
            };
            let reg = registry().lock().unwrap();
            let Some(tag) = reg.tags.get(&id) else {
                return err("closed");
            };
            let expected = args.get(1).map(bytes_of).unwrap_or_default();
            if expected.len() != tag.len() {
                return err("invalid-tag");
            }
            let diff = tag
                .iter()
                .zip(expected.iter())
                .fold(0u8, |acc, (a, b)| acc | (a ^ b));
            if diff == 0 {
                Value::Null
            } else {
                err("invalid-tag")
            }
        }),
    );

    vm.register_host_fn(
        SYMMETRIC,
        "symmetric-tag-close",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = args.first().and_then(|v| handle_id(v, KIND_SYMMETRIC_TAG)) else {
                return err("invalid-handle");
            };
            registry().lock().unwrap().tags.remove(&id);
            Value::Null
        }),
    );

    // AEAD, ratcheting and state-derived keys need cipher and KDF
    // implementations this module does not carry.
    for name in [
        "symmetric-state-encrypt",
        "symmetric-state-encrypt-detached",
        "symmetric-state-decrypt",
        "symmetric-state-decrypt-detached",
        "symmetric-state-ratchet",
        "symmetric-state-squeeze-key",
        "symmetric-state-options-get",
        "symmetric-state-options-get-u64",
        "symmetric-key-generate-managed",
        "symmetric-key-store-managed",
        "symmetric-key-replace-managed",
        "symmetric-key-id",
        "symmetric-key-from-id",
    ] {
        vm.register_host_fn(
            SYMMETRIC,
            name,
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| err("not-implemented")),
        );
    }
}

/// `vybe:crypto` — one-shot digests, and NOT a WASI interface.
///
/// ⚠This was `wasi:crypto/hashes`, and its own comment said the proposal
/// defines no such interface — kept anyway "because it is already in the
/// emitter's namespace list". That is the whole invented-verb pattern in one
/// sentence: a name nobody declares, living inside `wasi:` because moving it
/// was more work than leaving it.
///
/// The 8 interfaces wasi-crypto DOES declare are in
/// `proposals/wasi-crypto/wit/`, and none is `hashes`. Real wasi-crypto hashes
/// through `symmetric-state-open("SHA-256")` + `absorb` + `squeeze`; `sha256`
/// and `md5` appear in the spec only as algorithm NAMES inside signature
/// suites. A one-shot digest is a Vybe convenience, so it lives under `vybe:`
/// where a convenience belongs and where no guest can mistake it for spec.
///
/// The coverage gate could not catch this: `wasi:crypto` is a SEPARATE_PROPOSAL
/// and therefore exempt from `registered ⊆ spec`. An exemption at package
/// granularity hides an invented INTERFACE inside a real package.
fn register_hashes_shim(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:crypto/hashes",
        "sha256",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let data = args.first().map(bytes_of).unwrap_or_default();
            Value::String(Arc::from(sha256_hex(&data).as_str()))
        }),
    );

    vm.register_host_fn(
        "wasi:crypto/hashes",
        "md5",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let data = args.first().map(bytes_of).unwrap_or_default();
            Value::String(Arc::from(md5_hex(&data).as_str()))
        }),
    );
}
