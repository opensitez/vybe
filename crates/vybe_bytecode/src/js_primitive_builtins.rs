//! # `wasm:js-{number,boolean,undefined,symbol,bigint}`
//!
//! Stage-1 `js-primitive-builtins` WebAssembly proposal — **not merged**.
//! Spec: `proposals/js-primitive-builtins/proposals/js-primitive-builtins/Overview.md`
//!
//! Extends `wasm:js-string` to cover the remaining JS primitive types:
//! number, boolean, undefined, symbol, bigint.
//!
//! Every function that the spec marks as `trap()` on bad input calls
//! `ctx.throw_value(...)` so the VM raises a trap.

use std::sync::Arc;
use crate::{HostContext, VM, Value};

/// Raise a WASM trap from a host builtin (spec `trap()`), matching the
/// `wasm:js-string` convention.
fn trap(ctx: &mut HostContext, msg: &str) {
    ctx.throw_value(Value::String(Arc::from(msg)));
}

fn is_neg_zero(n: f64) -> bool {
    n == 0.0 && n.is_sign_negative()
}

pub fn register(vm: &mut VM) {
    register_number(vm);
    register_boolean(vm);
    register_undefined(vm);
    register_symbol(vm);
    register_bigint(vm);
}

// ── wasm:js-number ────────────────────────────────────────────────────

fn register_number(vm: &mut VM) {
    // test(externref) -> i32 — 1 if JS number
    vm.register_host_fn(
        "wasm:js-number",
        "test",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(match args.first() {
                Some(Value::F64(_)) | Some(Value::F32(_)) | Some(Value::I32(_))
                | Some(Value::I64(_)) => 1,
                _ => 0,
            })
        }),
    );

    // testI32(externref) -> i32 — 1 if integer-valued, in i32 range, not -0
    vm.register_host_fn(
        "wasm:js-number",
        "testI32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(match args.first() {
                Some(Value::I32(_)) => 1,
                Some(Value::F64(n)) => {
                    if is_neg_zero(*n) {
                        return Value::I32(0);
                    }
                    if (*n as i32) as f64 == *n { 1 } else { 0 }
                }
                Some(Value::F32(n)) => {
                    let n = *n as f64;
                    if is_neg_zero(n) {
                        return Value::I32(0);
                    }
                    if (n as i32) as f64 == n { 1 } else { 0 }
                }
                _ => 0,
            })
        }),
    );

    // testU32(externref) -> i32 — 1 if integer-valued, in [0, 2^32), not -0
    vm.register_host_fn(
        "wasm:js-number",
        "testU32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(match args.first() {
                Some(Value::I32(n)) => {
                    if *n >= 0 {
                        1
                    } else {
                        0
                    }
                }
                Some(Value::F64(n)) => {
                    if is_neg_zero(*n) {
                        return Value::I32(0);
                    }
                    if (*n as u32) as f64 == *n { 1 } else { 0 }
                }
                Some(Value::F32(n)) => {
                    let n = *n as f64;
                    if is_neg_zero(n) {
                        return Value::I32(0);
                    }
                    if (n as u32) as f64 == n { 1 } else { 0 }
                }
                _ => 0,
            })
        }),
    );

    // fromF64(f64) -> externref — identity box
    vm.register_host_fn(
        "wasm:js-number",
        "fromF64",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0))
        }),
    );

    // fromI32(i32) -> externref — box signed i32
    vm.register_host_fn(
        "wasm:js-number",
        "fromI32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(args.first().map(|v| v.as_i32()).unwrap_or(0))
        }),
    );

    // fromU32(i32) -> externref — reinterpret i32 bits as u32, return as f64
    vm.register_host_fn(
        "wasm:js-number",
        "fromU32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::F64((args.first().map(|v| v.as_i32()).unwrap_or(0) as u32) as f64)
        }),
    );

    // toF64(externref) -> f64 — trap if not a number
    vm.register_host_fn(
        "wasm:js-number",
        "toF64",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match args.first() {
            Some(Value::F64(n)) => Value::F64(*n),
            Some(Value::F32(n)) => Value::F64(*n as f64),
            Some(Value::I32(n)) => Value::F64(*n as f64),
            Some(Value::I64(n)) => Value::F64(*n as f64),
            _ => {
                ctx.throw_value(Value::String(Arc::from(
                    "TypeError: wasm:js-number.toF64 — not a number",
                )));
                Value::Null
            }
        }),
    );

    // toI32(externref) -> i32 — trap if not a number, not integer-valued, or -0
    vm.register_host_fn(
        "wasm:js-number",
        "toI32",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let n = match args.first() {
                Some(Value::I32(n)) => return Value::I32(*n),
                Some(Value::F64(n)) => *n,
                Some(Value::F32(n)) => *n as f64,
                Some(Value::I64(n)) => *n as f64,
                _ => {
                    ctx.throw_value(Value::String(Arc::from(
                        "TypeError: wasm:js-number.toI32 — not a number",
                    )));
                    return Value::Null;
                }
            };
            if is_neg_zero(n) || (n as i32) as f64 != n {
                ctx.throw_value(Value::String(Arc::from(
                    "TypeError: wasm:js-number.toI32 — not an i32 integer",
                )));
                return Value::Null;
            }
            Value::I32(n as i32)
        }),
    );

    // toU32(externref) -> i32 — trap if not a number, not u32-valued, or -0; returns bits as i32
    vm.register_host_fn(
        "wasm:js-number",
        "toU32",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let n = match args.first() {
                Some(Value::I32(n)) if *n >= 0 => return Value::I32(*n),
                Some(Value::F64(n)) => *n,
                Some(Value::I64(n)) if *n >= 0 && *n <= u32::MAX as i64 => {
                    return Value::I32(*n as i32);
                }
                _ => {
                    ctx.throw_value(Value::String(Arc::from(
                        "TypeError: wasm:js-number.toU32 — not a number",
                    )));
                    return Value::Null;
                }
            };
            if is_neg_zero(n) || (n as u32) as f64 != n {
                ctx.throw_value(Value::String(Arc::from(
                    "TypeError: wasm:js-number.toU32 — not a u32 integer",
                )));
                return Value::Null;
            }
            Value::I32(n as u32 as i32)
        }),
    );
}

// ── wasm:js-boolean ───────────────────────────────────────────────────

fn register_boolean(vm: &mut VM) {
    // test(externref) -> i32 — 1 if JS boolean
    vm.register_host_fn(
        "wasm:js-boolean",
        "test",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(if matches!(args.first(), Some(Value::Bool(_))) {
                1
            } else {
                0
            })
        }),
    );

    // cast(externref) -> i32 — extract as 0/1.
    // Spec (js-primitive-builtins, "wasm:js-boolean" "cast"):
    //     if (typeof x !== "boolean") trap();
    //     return x ? 1 : 0;
    // Traps on every non-boolean. This is safe for the emitter because `cast`
    // is only ever emitted inside an `if test(v)` guard (see
    // `vybe_emitter::ops::emit_dyn_to_bool`), and `test` returns 1 *only* for
    // Value::Bool — so a non-boolean never reaches this call.
    vm.register_host_fn(
        "wasm:js-boolean",
        "cast",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match args.first() {
            Some(Value::Bool(b)) => Value::I32(if *b { 1 } else { 0 }),
            _ => {
                trap(ctx, "TypeError: wasm:js-boolean.cast — not a boolean");
                Value::Null
            }
        }),
    );

    // fromI32(i32) -> externref — convert i32 to Bool value.
    // 0 → Bool(false), nonzero → Bool(true).
    vm.register_host_fn(
        "wasm:js-boolean",
        "fromI32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let v = match args.first() {
                Some(Value::I32(n)) => *n != 0,
                Some(Value::F64(n)) => *n != 0.0,
                Some(Value::Bool(b)) => *b,
                _ => false,
            };
            Value::Bool(v)
        }),
    );
}

// ── wasm:js-undefined ─────────────────────────────────────────────────

fn register_undefined(vm: &mut VM) {
    // test(externref) -> i32 — 1 if undefined
    vm.register_host_fn(
        "wasm:js-undefined",
        "test",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(if matches!(args.first(), Some(Value::Undefined)) {
                1
            } else {
                0
            })
        }),
    );
}

// ── wasm:js-symbol ────────────────────────────────────────────────────

fn register_symbol(vm: &mut VM) {
    // test(externref) -> i32 — 1 if JS symbol
    vm.register_host_fn(
        "wasm:js-symbol",
        "test",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(if matches!(args.first(), Some(Value::Symbol(_))) {
                1
            } else {
                0
            })
        }),
    );

    // equals(externref, externref) -> i32
    // Identity equality (Arc::ptr_eq). Traps if either arg is not a symbol or null.
    vm.register_host_fn(
        "wasm:js-symbol",
        "equals",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let x = args.first().unwrap_or(&Value::Null);
            let y = args.get(1).unwrap_or(&Value::Null);
            if !matches!(x, Value::Symbol(_) | Value::Null)
                || !matches!(y, Value::Symbol(_) | Value::Null)
            {
                ctx.throw_value(Value::String(Arc::from(
                    "TypeError: wasm:js-symbol.equals — not a symbol or null",
                )));
                return Value::Null;
            }
            Value::I32(match (x, y) {
                (Value::Symbol(a), Value::Symbol(b)) => {
                    if Arc::ptr_eq(a, b) {
                        1
                    } else {
                        0
                    }
                }
                (Value::Null, Value::Null) => 1,
                _ => 0,
            })
        }),
    );
}

// ── wasm:js-bigint ────────────────────────────────────────────────────

fn register_bigint(vm: &mut VM) {
    // test(externref) -> i32 — 1 if JS bigint
    vm.register_host_fn(
        "wasm:js-bigint",
        "test",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(if matches!(args.first(), Some(Value::BigInt(_))) {
                1
            } else {
                0
            })
        }),
    );
    // fromI64(i64) -> externref — wrap an i64 as a BigInt value
    vm.register_host_fn(
        "wasm:js-bigint",
        "fromI64",
        Box::new(
            |_ctx: &mut HostContext, args: &[Value]| match args.first() {
                Some(Value::I64(v)) => Value::bigint_i64(*v),
                Some(Value::I32(v)) => Value::bigint_i64(*v as i64),
                Some(Value::F64(v)) => Value::bigint_i64(*v as i64),
                Some(Value::BigInt(v)) => Value::BigInt(v.clone()),
                _ => Value::bigint_i64(0),
            },
        ),
    );
}
