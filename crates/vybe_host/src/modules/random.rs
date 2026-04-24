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
//! Vybe-convenience (NOT in WASI) lives under `vybe:random/*`:
//!   * `vybe:random/random` — Math.random()-style f64 in [0,1)
//!   * `vybe:random/randomInt(min, max)` — inclusive-range integer
//!   * `vybe:random/uuid` — UUID v4 string
//!
//! [`wasi-random`/random.wit]: proposals/WASI/proposals/random/wit/random.wit

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

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
    vm.register_host_fn("wasi:random/random", "get-random-bytes", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
        random_bytes_value(n)
    }));
    vm.register_host_fn("wasi:random/random", "get-random-u64", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(next_u64() as f64)
    }));

    // ── wasi:random/insecure ───────────────────────────────────────────
    // Spec: interface `insecure` at `wasi:random/insecure`.
    //   get-insecure-random-bytes: func(len: u64) -> list<u8>
    //   get-insecure-random-u64:   func() -> u64
    vm.register_host_fn("wasi:random/insecure", "get-insecure-random-bytes", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
        random_bytes_value(n)
    }));
    vm.register_host_fn("wasi:random/insecure", "get-insecure-random-u64", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(next_u64() as f64)
    }));

    // ── wasi:random/insecure-seed ──────────────────────────────────────
    // Spec: interface `insecure-seed` at `wasi:random/insecure-seed`.
    //   insecure-seed: func() -> tuple<u64, u64>
    // The two u64s form a 128-bit seed value. MVP returns the xorshift
    // state and a derivative.
    vm.register_host_fn("wasi:random/insecure-seed", "insecure-seed", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let a = next_u64();
        let b = next_u64();
        let pair = vec![Value::F64(a as f64), Value::F64(b as f64)];
        Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
    }));

    // ── vybe:random/* — conveniences NOT in WASI ──────────────────────
    //
    // `Math.random()`-style float in [0, 1). WASI's `random/random`
    // gives raw u64; ECMA-262 `Math.random` gives a float. Different
    // concepts, so this lives under `vybe:random`.
    vm.register_host_fn("vybe:random", "random", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(next_f64())
    }));

    // Random integer in inclusive range. WASI gives raw bytes/u64;
    // mapping to a range is Vybe-convenience.
    vm.register_host_fn("vybe:random", "randomInt", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let min = args.first().map(|v| v.as_f64() as i64).unwrap_or(0);
        let max = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(100);
        if max <= min { return Value::F64(min as f64); }
        let range = (max - min + 1) as u64;
        let val = min + (next_u64() % range) as i64;
        Value::F64(val as f64)
    }));

    // UUID v4 — not a WASI concept. Generates over local randomness.
    vm.register_host_fn("vybe:random", "uuid", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let a = next_u64();
        let b = next_u64();
        let s = format!(
            "{:08x}-{:04x}-4{:03x}-{:04x}-{:012x}",
            (a >> 32) as u32,
            (a >> 16) as u16 & 0xFFFF,
            a as u16 & 0x0FFF,
            (b >> 48) as u16 & 0x3FFF | 0x8000,
            b & 0xFFFFFFFFFFFF,
        );
        Value::String(Arc::from(s.as_str()))
    }));
}
