//! WASI random implementation.
//!
//! Split across:
//!   * `wasi:random/random` — cryptographically-strong randomness per
//!     [`wasi-random`/random.wit]. `get-random-bytes: func(len: u64) -> list<u8>`
//!     and `get-random-u64: func() -> u64`.
//!   * `wasi:random/insecure` — fast pseudo-random, not for security.
//!     `get-insecure-random-bytes: func(len: u64) -> list<u8>` and
//!     `get-insecure-random-u64: func() -> u64`.
//!   * `wasi:random/insecure-seed` — DoS-resistance hash seed.
//!     `insecure-seed: func() -> tuple<u64, u64>`.
//!
//! `wasi:random/random` is backed by the operating system CSPRNG
//! (`getrandom`, i.e. `getrandom(2)`/`/dev/urandom` on Unix and
//! `BCryptGenRandom` on Windows), because the interface's contract is
//! normative:
//!
//! > must produce data at least as cryptographically secure and fast as an
//! > adequately seeded cryptographically-secure pseudo-random number
//! > generator (CSPRNG). It must not block … including on the first request
//! > … The returned data must always be unpredictable.
//!
//! The xorshift64 generator below therefore serves ONLY `wasi:random/insecure`
//! and `wasi:random/insecure-seed`, whose contracts explicitly disclaim
//! cryptographic strength ("There are no requirements on the values of the
//! returned bytes"). Routing the secure interface through xorshift — as this
//! module used to — is a security defect, not a fidelity gap: xorshift64 is
//! trivially invertible, so one observed output reveals the whole state and
//! every past and future value with it.
//!
//! Numbers are carried as `Value::F64`, the platform's numeric
//! representation (same as `wasi:clocks`). That costs the low 11 bits of a
//! `u64`, so `get-random-bytes` — not `get-random-u64` — is the full-entropy
//! path. The carrier is load-bearing elsewhere: PHP's `rand()` scales this
//! value by `r / 2^64` in f64 (`php/src/emitter/numeric_adapter.rs`).
//!
//! Vybe-convenience extensions live under the same `wasi:random/random`
//! namespace, on top of the WASI primitives:
//!   * `wasi:random/random.random` — Math.random()-style f64 in [0,1)
//!   * `wasi:random/random.randomInt(min, max)` — inclusive-range integer
//!   * `wasi:random/random.uuid` — UUID v4 string
//! These names aren't in the WASI proposal; they're convenience helpers
//! callers reach via `import * as random from "wasi:random/random"`.
//!
//! [`wasi-random`/random.wit]: proposals/WASI/proposals/random/wit/random.wit

use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

/// OS entropy for the CSPRNG-grade interface. `None` when the platform has
/// no entropy source — callers must surface that rather than silently
/// degrade to a predictable generator.
///
/// `getrandom` never blocks after the system pool is initialised, which is
/// what the interface's "must not block … including on the first request"
/// requires.
pub fn secure_bytes(n: usize) -> Option<Vec<u8>> {
    let mut buf = vec![0u8; n];
    getrandom::getrandom(&mut buf).ok()?;
    Some(buf)
}

/// A single CSPRNG `u64`.
fn secure_u64() -> Option<u64> {
    let bytes = secure_bytes(8)?;
    let mut arr = [0u8; 8];
    arr.copy_from_slice(&bytes);
    Some(u64::from_le_bytes(arr))
}

// Simple xorshift64 PRNG state — thread-local for safety.
// Backs ONLY `wasi:random/insecure` and `wasi:random/insecure-seed`, whose
// contracts disclaim cryptographic strength. Never the secure interface.
thread_local! {
    static RNG_STATE: std::cell::RefCell<u64> = std::cell::RefCell::new(
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos() as u64
    );
}

fn next_u64() -> u64 {
    RNG_STATE.with(|state| {
        let mut s = state.borrow_mut();
        // xorshift64
        *s ^= *s << 13;
        *s ^= *s >> 7;
        *s ^= *s << 17;
        *s
    })
}

fn next_f64() -> f64 {
    (next_u64() >> 11) as f64 / (1u64 << 53) as f64
}

/// `tuple<u64, u64>` — the shape both `insecure-seed` (0.2) and
/// `get-insecure-seed` (0.3) return.
fn insecure_seed_pair() -> Value {
    let a = next_u64();
    let b = next_u64();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::F64(a as f64),
        Value::F64(b as f64),
    ])))
}

fn bytes_to_list(bytes: &[u8]) -> Value {
    let values: Vec<Value> = bytes.iter().map(|b| Value::F64(*b as f64)).collect();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

/// Insecure `list<u8>` from the xorshift stream.
fn insecure_bytes_value(n: usize) -> Value {
    let bytes: Vec<u8> = (0..n).map(|_| (next_u64() & 0xFF) as u8).collect();
    bytes_to_list(&bytes)
}

/// CSPRNG `list<u8>`. Falls back to nothing — an empty list is a visible
/// failure, whereas substituting the insecure stream would hand back
/// predictable bytes under a name that promises unpredictability.
fn secure_bytes_value(n: usize) -> Value {
    match secure_bytes(n) {
        Some(bytes) => bytes_to_list(&bytes),
        None => bytes_to_list(&[]) }
}

pub fn register(vm: &mut VM) {
    // ── wasi:random/random ─────────────────────────────────────────────
    // Spec: interface `random` at `wasi:random/random`.
    //   get-random-bytes: func(len: u64) -> list<u8>
    //   get-random-u64:   func() -> u64
    vm.register_host_fn(
        "wasi:random/random",
        "get-random-bytes",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args
                .first()
                .map(|v| v.as_f64())
                .filter(|len| *len > 0.0)
                .map(|len| len as usize)
                .unwrap_or(0);
            secure_bytes_value(n)
        }),
    );
    vm.register_host_fn(
        "wasi:random/random",
        "get-random-u64",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::F64(secure_u64().unwrap_or(0) as f64)
        }),
    );

    // ── wasi:random/insecure ───────────────────────────────────────────
    // Spec: interface `insecure` at `wasi:random/insecure`.
    //   get-insecure-random-bytes: func(len: u64) -> list<u8>
    //   get-insecure-random-u64:   func() -> u64
    vm.register_host_fn(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args
                .first()
                .map(|v| v.as_f64())
                .filter(|len| *len > 0.0)
                .map(|len| len as usize)
                .unwrap_or(0);
            insecure_bytes_value(n)
        }),
    );
    vm.register_host_fn(
        "wasi:random/insecure",
        "get-insecure-random-u64",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(next_u64() as f64)),
    );

    // ── wasi:random/insecure-seed ──────────────────────────────────────
    // Spec: interface `insecure-seed` at `wasi:random/insecure-seed`.
    //   insecure-seed: func() -> tuple<u64, u64>
    // The two u64s form a 128-bit seed value. MVP returns the xorshift
    // state and a derivative.
    vm.register_host_fn(
        "wasi:random/insecure-seed",
        "insecure-seed",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| insecure_seed_pair()),
    );

    // 0.3.0 renamed the function to `get-insecure-seed`
    // (`proposals/random/wit/insecure-seed.wit`); the 0.2 name stays bound so
    // components built against either revision resolve.
    vm.register_host_fn(
        "wasi:random/insecure-seed",
        "get-insecure-seed",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| insecure_seed_pair()),
    );

    // ── Convenience extensions under wasi:random/random ───────────────
    vm.register_host_fn(
        "wasi:random/random",
        "random",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(next_f64())),
    );
    vm.register_host_fn(
        "wasi:random/random",
        "randomInt",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let min = args.first().map(|v| v.as_f64() as i64).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(min);
            if min >= max {
                return Value::F64(min as f64);
            }
            let range = (max - min + 1) as u64;
            let r = (next_u64() % range) as i64 + min;
            Value::F64(r as f64)
        }),
    );
    vm.register_host_fn(
        "wasi:random/random",
        "uuid",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let a = next_u64();
            let b = next_u64();
            let bytes: [u8; 16] = {
                let mut buf = [0u8; 16];
                buf[..8].copy_from_slice(&a.to_le_bytes());
                buf[8..].copy_from_slice(&b.to_le_bytes());
                buf[6] = (buf[6] & 0x0f) | 0x40; // version 4
                buf[8] = (buf[8] & 0x3f) | 0x80; // variant 1
                buf
            };
            let s = format!(
                "{:02x}{:02x}{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
                bytes[0], bytes[1], bytes[2], bytes[3],
                bytes[4], bytes[5], bytes[6], bytes[7],
                bytes[8], bytes[9], bytes[10], bytes[11],
                bytes[12], bytes[13], bytes[14], bytes[15],
            );
            Value::String(Arc::from(s.as_str()))
        }),
    );
}
