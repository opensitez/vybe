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
//! The `u64` results use an `I64` carrier to preserve all 64 bits. Consumers
//! that need floating-point scaling convert it explicitly as unsigned.
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

use vybe_runtime::value::Object;
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType, Value};

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
    let mut bytes = [0u8; 8];
    getrandom::getrandom(&mut bytes).ok()?;
    Some(u64::from_le_bytes(bytes))
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
        advance(&mut *s)
    })
}

fn advance(state: &mut u64) -> u64 {
    *state ^= *state << 13;
    *state ^= *state >> 7;
    *state ^= *state << 17;
    *state
}

/// `tuple<u64, u64>` — the shape `get-insecure-seed` returns.
fn insecure_seed_pair() -> Value {
    let a = next_u64();
    let b = next_u64();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I64(a as i64),
        Value::I64(b as i64),
    ])))
}

fn bytes_to_list(bytes: &[u8]) -> Value {
    let values: Vec<Value> = bytes.iter().map(|b| Value::F64(*b as f64)).collect();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

/// Insecure `list<u8>` from the xorshift stream.
fn insecure_bytes_value(n: usize) -> Value {
    let values = RNG_STATE.with(|state| {
        let mut state = state.borrow_mut();
        (0..n)
            .map(|_| Value::F64((advance(&mut *state) & 0xff) as f64))
            .collect()
    });
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

#[cfg(test)]
mod random_tests {
    use super::{RNG_STATE, advance, insecure_bytes_value, insecure_seed_pair, next_u64, secure_u64};
    use vybe_runtime::value::ObjectKind;
    use vybe_runtime::Value;

    #[test]
    fn bulk_insecure_bytes_keep_the_rng_sequence() {
        let original = RNG_STATE.with(|state| state.replace(0x1234_5678_9abc_def0));
        let mut expected_state = 0x1234_5678_9abc_def0;
        let expected: Vec<Value> = (0..32)
            .map(|_| Value::F64((advance(&mut expected_state) & 0xff) as f64))
            .collect();
        let Value::Object(array) = insecure_bytes_value(32) else {
            panic!("expected byte list")
        };
        let array = array.lock().unwrap();
        let ObjectKind::Array(actual) = &array.kind else {
            panic!("expected array")
        };
        assert_eq!(actual, &expected);
        drop(array);
        assert_eq!(next_u64(), advance(&mut expected_state));
        RNG_STATE.with(|state| *state.borrow_mut() = original);
    }

    #[test]
    fn secure_u64_reads_eight_entropy_bytes() {
        assert!(secure_u64().is_some());
    }

    #[test]
    fn insecure_seed_preserves_all_u64_bits() {
        let original = RNG_STATE.with(|state| state.replace(0xfedc_ba98_7654_3210));
        let mut expected = 0xfedc_ba98_7654_3210;
        let first = advance(&mut expected);
        let second = advance(&mut expected);
        let pair = insecure_seed_pair();
        let Value::Object(array) = pair else {
            panic!("expected seed pair")
        };
        let array = array.lock().unwrap();
        let ObjectKind::Array(values) = &array.kind else {
            panic!("expected seed array")
        };
        assert_eq!(
            values,
            &vec![Value::I64(first as i64), Value::I64(second as i64)]
        );
        drop(array);
        RNG_STATE.with(|state| *state.borrow_mut() = original);
    }

    #[test]
    #[ignore = "local secure-u64 allocation benchmark"]
    fn bench_secure_u64_stack_against_heap() {
        use std::time::Instant;

        for _ in 0..3 {
            let start = Instant::now();
            for _ in 0..100_000 {
                std::hint::black_box(secure_u64().unwrap());
            }
            let stack = start.elapsed();

            let start = Instant::now();
            for _ in 0..100_000 {
                let bytes = super::secure_bytes(8).unwrap();
                std::hint::black_box(u64::from_le_bytes(bytes.try_into().unwrap()));
            }
            eprintln!("secure u64: stack {stack:?}, heap {:?}", start.elapsed());
        }
    }
}

/// CSPRNG `list<u8>`. Falls back to nothing — an empty list is a visible
/// failure, whereas substituting the insecure stream would hand back
/// predictable bytes under a name that promises unpredictability.
fn secure_bytes_value(n: usize) -> Value {
    match secure_bytes(n) {
        Some(bytes) => bytes_to_list(&bytes),
        None => bytes_to_list(&[]),
    }
}

/// Declare a `wasi:random/*` function.
///
/// None of these three interfaces owns a resource: randomness is produced, not
/// held, so every function is free and returns plain data. The registered names
/// are already the WIT spelling, so the signature name is the name.
fn random_fn(
    vm: &mut VM,
    module: &str,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(HostFnDecl::new(module, name, call).with_sig(FuncSig {
        name: name.to_string(),
        params: Param::unnamed_list(params),
        results,
    }));
}

/// `list<u8>` — a byte list. `ValType` has no u8, and `I32` is the narrowest
/// integer it carries, so that is what a byte declares as.
fn byte_list() -> ValType {
    ValType::List(Box::new(ValType::I32))
}

/// `tuple<u64, u64>` — the Component Model spells a tuple as a record with
/// positional field names.
fn u64_pair() -> ValType {
    ValType::Record(vec![
        ("0".to_string(), ValType::I64),
        ("1".to_string(), ValType::I64),
    ])
}

pub fn register(vm: &mut VM) {
    // ── wasi:random/random ─────────────────────────────────────────────
    // Spec: interface `random` at `wasi:random/random`.
    //   get-random-bytes: func(len: u64) -> list<u8>
    //   get-random-u64:   func() -> u64
    random_fn(
        vm,
        "wasi:random/random",
        "get-random-bytes",
        vec![ValType::I64],
        vec![byte_list()],
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
    random_fn(
        vm,
        "wasi:random/random",
        "get-random-u64",
        vec![],
        vec![ValType::I64],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::I64(secure_u64().unwrap_or(0) as i64)
        }),
    );

    // ── wasi:random/insecure ───────────────────────────────────────────
    // Spec: interface `insecure` at `wasi:random/insecure`.
    //   get-insecure-random-bytes: func(len: u64) -> list<u8>
    //   get-insecure-random-u64:   func() -> u64
    random_fn(
        vm,
        "wasi:random/insecure",
        "get-insecure-random-bytes",
        vec![ValType::I64],
        vec![byte_list()],
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
    random_fn(
        vm,
        "wasi:random/insecure",
        "get-insecure-random-u64",
        vec![],
        vec![ValType::I64],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::I64(next_u64() as i64)),
    );

    // ── wasi:random/insecure-seed ──────────────────────────────────────
    // Spec: interface `insecure-seed` at `wasi:random/insecure-seed`.
    //   insecure-seed: func() -> tuple<u64, u64>
    // The two u64s form a 128-bit seed value. MVP returns the xorshift
    // state and a derivative.
    // 0.3 renamed this function from `insecure-seed` to `get-insecure-seed`
    // (`proposals/WASI/proposals/random/wit/insecure-seed.wit`). Both were
    // bound until 2026-08-21, "so components built against either revision
    // resolve" — which is the reason a deleted name never leaves.
    random_fn(
        vm,
        "wasi:random/insecure-seed",
        "get-insecure-seed",
        vec![],
        vec![u64_pair()],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| insecure_seed_pair()),
    );

    // `random`, `randomInt` and `uuid` USED TO BE REGISTERED HERE, under the
    // banner "Convenience extensions".
    //
    // `wasi:random@0.3.1` declares `get-random-bytes` and `get-random-u64`,
    // and nothing else. These three were invented and put in the `wasi:`
    // namespace anyway, so a guest built against the WIT could not call them
    // and a conforming runtime could not answer them — the prefix was a claim
    // this module did not keep.
    //
    // `random` and `randomInt` had no caller in the tree at all. `uuid` had
    // exactly one — Python's `tempfile` name — and it never needed a canonical
    // UUID, only a unique token; it composes one from `get-random-u64`.
}
