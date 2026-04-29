//! ECMA-262 §25.4 — Atomics.
//!
//!   §25.4.1  Atomics.add(typedArray, index, value) → previous value
//!   §25.4.2  Atomics.and(typedArray, index, value)
//!   §25.4.3  Atomics.compareExchange(typedArray, index, expected, replacement)
//!   §25.4.4  Atomics.exchange(typedArray, index, value)
//!   §25.4.5  Atomics.isLockFree(size) → bool
//!   §25.4.6  Atomics.load(typedArray, index) → value
//!   §25.4.7  Atomics.notify(typedArray, index, count) → woken
//!   §25.4.8  Atomics.or(typedArray, index, value)
//!   §25.4.9  Atomics.store(typedArray, index, value) → value
//!   §25.4.10 Atomics.sub(typedArray, index, value)
//!   §25.4.11 Atomics.wait(typedArray, index, value, timeout) → "ok"|"not-equal"|"timed-out"
//!   §25.4.12 Atomics.waitAsync — Stage-4 proposal
//!   §25.4.13 Atomics.xor(typedArray, index, value)
//!
//! These are adapters mirroring the WASM `*.atomic.rmw.*` /
//! `atomic.fence` / `memory.atomic.{wait,notify}` opcodes — the
//! compiler emits opcodes directly when the operand is a known
//! TypedArray view; this module is the dynamic-dispatch fallback.
//!
//! Vybe's atomics fall back to non-atomic Vec ops in the MVP impl —
//! true sequential consistency requires SharedArrayBuffer backed by
//! `Arc<Mutex<Vec<u8>>>`. The data type is correct (see
//! `crate::ecma::arraybuffer`); the host fns route the read/write
//! through the same buffer the underlying TypedArray points at.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::ObjectKind;

/// Resolve a typed-array argument to its backing buffer + offset metadata.
/// Returns (buffer, byteOffset, elementByteLength) when valid.
fn typed_array_buffer(args: &[Value], idx: usize) -> Option<(Arc<Mutex<Vec<u8>>>, usize, usize)> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        let buffer = o.properties.get("buffer")?;
        let byte_offset = o.properties.get("byteOffset").map(|v| v.as_f64() as usize).unwrap_or(0);
        let bpe = o.properties.get("BYTES_PER_ELEMENT").map(|v| v.as_f64() as usize).unwrap_or(4);
        if let Value::Object(buf_obj) = buffer {
            let bo = buf_obj.lock().unwrap();
            if let ObjectKind::ArrayBuffer(ref ab) = bo.kind {
                return Some((ab.bytes.clone(), byte_offset, bpe));
            }
        }
    }
    None
}

/// Read a value from the typed-array backing buffer at the given element index.
fn atomic_load(buf: &Arc<Mutex<Vec<u8>>>, byte_offset: usize, idx: usize, bpe: usize) -> i64 {
    let data = buf.lock().unwrap();
    let off = byte_offset + idx * bpe;
    if off + bpe > data.len() { return 0; }
    match bpe {
        1 => data[off] as i64,
        2 => i16::from_le_bytes([data[off], data[off+1]]) as i64,
        4 => i32::from_le_bytes([data[off], data[off+1], data[off+2], data[off+3]]) as i64,
        8 => i64::from_le_bytes([
            data[off], data[off+1], data[off+2], data[off+3],
            data[off+4], data[off+5], data[off+6], data[off+7],
        ]),
        _ => 0,
    }
}

fn atomic_store(buf: &Arc<Mutex<Vec<u8>>>, byte_offset: usize, idx: usize, bpe: usize, val: i64) {
    let mut data = buf.lock().unwrap();
    let off = byte_offset + idx * bpe;
    if off + bpe > data.len() { return; }
    match bpe {
        1 => data[off] = val as u8,
        2 => {
            let bytes = (val as i16).to_le_bytes();
            data[off..off+2].copy_from_slice(&bytes);
        }
        4 => {
            let bytes = (val as i32).to_le_bytes();
            data[off..off+4].copy_from_slice(&bytes);
        }
        8 => {
            let bytes = val.to_le_bytes();
            data[off..off+8].copy_from_slice(&bytes);
        }
        _ => {}
    }
}

pub fn register(vm: &mut VM) {
    macro_rules! rmw {
        ($name:expr, $op:expr) => {
            vm.register_host_fn("ecma:atomics", $name, Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
                let val = args.get(2).map(|v| v.as_f64() as i64).unwrap_or(0);
                if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                    let prev = atomic_load(&buf, off, idx, bpe);
                    let new_val: i64 = $op(prev, val);
                    atomic_store(&buf, off, idx, bpe, new_val);
                    return Value::F64(prev as f64);
                }
                Value::F64(0.0)
            }));
        };
    }

    rmw!("add", |a: i64, b: i64| a.wrapping_add(b));
    rmw!("sub", |a: i64, b: i64| a.wrapping_sub(b));
    rmw!("and", |a: i64, b: i64| a & b);
    rmw!("or",  |a: i64, b: i64| a | b);
    rmw!("xor", |a: i64, b: i64| a ^ b);

    // exchange(ta, idx, value) — store value, return previous.
    vm.register_host_fn("ecma:atomics", "exchange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let val = args.get(2).map(|v| v.as_f64() as i64).unwrap_or(0);
        if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
            let prev = atomic_load(&buf, off, idx, bpe);
            atomic_store(&buf, off, idx, bpe, val);
            return Value::F64(prev as f64);
        }
        Value::F64(0.0)
    }));

    // compareExchange(ta, idx, expected, replacement)
    vm.register_host_fn("ecma:atomics", "compareExchange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let expected = args.get(2).map(|v| v.as_f64() as i64).unwrap_or(0);
        let replacement = args.get(3).map(|v| v.as_f64() as i64).unwrap_or(0);
        if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
            let prev = atomic_load(&buf, off, idx, bpe);
            if prev == expected {
                atomic_store(&buf, off, idx, bpe, replacement);
            }
            return Value::F64(prev as f64);
        }
        Value::F64(0.0)
    }));

    vm.register_host_fn("ecma:atomics", "load", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
            return Value::F64(atomic_load(&buf, off, idx, bpe) as f64);
        }
        Value::F64(0.0)
    }));

    vm.register_host_fn("ecma:atomics", "store", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let val = args.get(2).map(|v| v.as_f64() as i64).unwrap_or(0);
        if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
            atomic_store(&buf, off, idx, bpe, val);
        }
        Value::F64(val as f64)
    }));

    // isLockFree(size) — int sizes (1,2,4,8) are lock-free on most arch.
    vm.register_host_fn("ecma:atomics", "isLockFree", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let size = args.first().map(|v| v.as_f64() as i32).unwrap_or(0);
        Value::Bool(matches!(size, 1 | 2 | 4 | 8))
    }));

    // wait(ta, idx, value, timeout?) → "ok" | "not-equal" | "timed-out".
    //
    // MVP: blocking wait isn't safe in this VM (no thread-park primitive
    // wired through HostContext). Returns "not-equal" if values differ,
    // "ok" otherwise after a tiny spin-yield. Real sequential consistency
    // ships when the SharedArrayBuffer Mutex grows a Condvar.
    vm.register_host_fn("ecma:atomics", "wait", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let expected = args.get(2).map(|v| v.as_f64() as i64).unwrap_or(0);
        if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
            let actual = atomic_load(&buf, off, idx, bpe);
            if actual != expected {
                return Value::String(Arc::from("not-equal"));
            }
        }
        Value::String(Arc::from("ok"))
    }));

    // notify(ta, idx, count?) → number of agents woken.
    // Always 0 in MVP (no waiters list).
    vm.register_host_fn("ecma:atomics", "notify", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(0.0)
    }));
}
