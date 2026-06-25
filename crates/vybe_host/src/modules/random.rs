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
//! The MVP implementation uses xorshift64 for all paths — not actually
//! cryptographically strong. When a real CSPRNG backs the host, the
//! `wasi:random/random` functions switch to it without any caller
//! change.
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

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

// Simple xorshift64 PRNG state — thread-local for safety.
// MVP backing for all random paths (including the nominally-secure
// `wasi:random/random`). Replace with a real CSPRNG when the host
// gains access to one.
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

/// Build a `list<u8>` as a Vybe Array object, matching what real WASI
/// `list<u8>` lowers to in our Value representation.
fn random_bytes_value(n: usize) -> Value {
    let mut bytes = Vec::with_capacity(n);
    for _ in 0..n {
        bytes.push(Value::F64((next_u64() & 0xFF) as f64));
    }
    Value::Object(Arc::new(Mutex::new(Object::new_array(bytes))))
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
            let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            random_bytes_value(n)
        }),
    );
    vm.register_host_fn(
        "wasi:random/random",
        "get-random-u64",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(next_u64() as f64)),
    );

    // ── wasi:random/insecure ───────────────────────────────────────────
    // Spec: interface `insecure` at `wasi:random/insecure`.
    //   get-insecure-random-bytes: func(len: u64) -> list<u8>
    //   get-insecure-random-u64:   func() -> u64
    vm.register_host_fn(
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            random_bytes_value(n)
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
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let a = next_u64();
            let b = next_u64();
            let pair = vec![Value::F64(a as f64), Value::F64(b as f64)];
            Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
        }),
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
