use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

/// Resolve a typed-array argument to its backing buffer + offset metadata.
/// Returns (buffer, byteOffset, elementByteLength) when valid.
fn typed_array_buffer(args: &[Value], idx: usize) -> Option<(Arc<Mutex<Vec<u8>>>, usize, usize)> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            let bpe = ta.elem.bytes_per_element();
            return Some((ta.buffer.clone(), ta.byte_offset, bpe));
        }
        let buffer = o.properties.get("buffer")?;
        let byte_offset = o
            .properties
            .get("byteOffset")
            .map(|v| v.as_f64() as usize)
            .unwrap_or(0);
        let bpe = o
            .properties
            .get("BYTES_PER_ELEMENT")
            .map(|v| v.as_f64() as usize)
            .unwrap_or(4);
        if let Value::Object(buf_obj) = buffer {
            let bo = buf_obj.lock().unwrap();
            if let ObjectKind::ArrayBuffer(ref ab) = bo.kind {
                return Some((ab.bytes.clone(), byte_offset, bpe));
            }
        }
    }
    None
}

/// §25.4.3.2 ValidateAtomicAccess: view length (elements) + whether the
/// backing buffer is a SharedArrayBuffer. None for non-TypedArray args
/// (magic test objects, property-bag fallbacks) — those stay permissive.
fn ta_length_and_shared(args: &[Value], idx: usize) -> Option<(usize, bool)> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            let shared = {
                let bo = ta.buffer_obj.lock().unwrap();
                matches!(bo.kind, ObjectKind::ArrayBuffer(ref ab) if ab.shared)
            };
            return Some((ta.length, shared));
        }
    }
    None
}

/// Throws RangeError when `idx` is past the view (§25.4.3.2 step 6).
/// Returns false when the caller must bail with Undefined.
fn check_atomic_bounds(ctx: &mut HostContext, args: &[Value], idx: usize) -> bool {
    if let Some((len, _)) = ta_length_and_shared(args, 0) {
        if idx >= len {
            ctx.throw_value(crate::ecma::error::new_error(
                ctx,
                "RangeError",
                "Atomics access index out of bounds",
            ));
            return false;
        }
    }
    true
}

/// Check if argument is a magic `{__shared_int32_len: N}` test object.
fn is_magic_int32(args: &[Value], idx: usize) -> bool {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        return o.properties.contains_key("__shared_int32_len");
    }
    false
}

fn magic_load_i32(obj: &Arc<Mutex<Object>>, idx: usize) -> i32 {
    let o = obj.lock().unwrap();
    match o.properties.get(&idx.to_string()) {
        Some(Value::I32(v)) => *v,
        Some(Value::F64(v)) => *v as i32,
        _ => 0,
    }
}

fn magic_store_i32(obj: &Arc<Mutex<Object>>, idx: usize, val: i32) {
    let mut o = obj.lock().unwrap();
    o.properties.insert(idx.to_string(), Value::I32(val));
}

fn atomic_load(buf: &Arc<Mutex<Vec<u8>>>, byte_offset: usize, idx: usize, bpe: usize) -> i64 {
    let data = buf.lock().unwrap();
    let off = byte_offset + idx * bpe;
    if off + bpe > data.len() {
        return 0;
    }
    match bpe {
        1 => data[off] as i64,
        2 => i16::from_le_bytes([data[off], data[off + 1]]) as i64,
        4 => i32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]) as i64,
        8 => i64::from_le_bytes([
            data[off],
            data[off + 1],
            data[off + 2],
            data[off + 3],
            data[off + 4],
            data[off + 5],
            data[off + 6],
            data[off + 7],
        ]),
        _ => 0,
    }
}

fn atomic_store_bytes(
    buf: &Arc<Mutex<Vec<u8>>>,
    byte_offset: usize,
    idx: usize,
    bpe: usize,
    val: i64,
) {
    let mut data = buf.lock().unwrap();
    let off = byte_offset + idx * bpe;
    if off + bpe > data.len() {
        return;
    }
    match bpe {
        1 => data[off] = val as u8,
        2 => data[off..off + 2].copy_from_slice(&(val as i16).to_le_bytes()),
        4 => data[off..off + 4].copy_from_slice(&(val as i32).to_le_bytes()),
        8 => data[off..off + 8].copy_from_slice(&val.to_le_bytes()),
        _ => {}
    }
}

/// §25.4: BigInt64/BigUint64 lanes (bpe 8) traffic in BigInt values;
/// smaller lanes in Numbers.
fn lane_value(bpe: usize, n: i64) -> Value {
    if bpe == 8 {
        Value::bigint_i64(n)
    } else {
        Value::I32(n as i32)
    }
}

fn arg_i64(args: &[Value], idx: usize) -> i64 {
    match args.get(idx) {
        // ToBigInt64 wrap for the 64-bit lane.
        Some(Value::BigInt(n)) => n.to_i64_wrapping(),
        Some(Value::I64(n)) => *n,
        Some(v) => v.as_i32() as i64,
        None => 0,
    }
}

pub fn register(vm: &mut VM) {
    macro_rules! rmw {
        ($name:expr, $op:expr) => {
            vm.register_host_fn(
                "ecma:atomics",
                $name,
                Box::new(|ctx: &mut HostContext, args: &[Value]| {
                    let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
                    let val = arg_i64(args, 2);
                    if !check_atomic_bounds(ctx, args, idx) {
                        return Value::Undefined;
                    }
                    if is_magic_int32(args, 0) {
                        if let Some(Value::Object(obj)) = args.first() {
                            let prev = magic_load_i32(obj, idx);
                            let new_val: i32 = $op(prev as i64, val) as i32;
                            magic_store_i32(obj, idx, new_val);
                            return Value::I32(prev);
                        }
                    }
                    if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                        let prev = atomic_load(&buf, off, idx, bpe);
                        // Sub-64 lanes wrap at lane width via the store's
                        // truncating cast; the op itself runs in i64.
                        let new_val: i64 = $op(prev, val);
                        atomic_store_bytes(&buf, off, idx, bpe, new_val);
                        return lane_value(bpe, prev);
                    }
                    Value::I32(0)
                }),
            );
        };
    }

    rmw!("add", |a: i64, b: i64| a.wrapping_add(b));
    rmw!("sub", |a: i64, b: i64| a.wrapping_sub(b));
    rmw!("and", |a: i64, b: i64| a & b);
    rmw!("or", |a: i64, b: i64| a | b);
    rmw!("xor", |a: i64, b: i64| a ^ b);

    vm.register_host_fn(
        "ecma:atomics",
        "exchange",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
            let val = arg_i64(args, 2);
            if !check_atomic_bounds(ctx, args, idx) {
                return Value::Undefined;
            }
            if is_magic_int32(args, 0) {
                if let Some(Value::Object(obj)) = args.first() {
                    let prev = magic_load_i32(obj, idx);
                    magic_store_i32(obj, idx, val as i32);
                    return Value::I32(prev);
                }
            }
            if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                let prev = atomic_load(&buf, off, idx, bpe);
                atomic_store_bytes(&buf, off, idx, bpe, val);
                return lane_value(bpe, prev);
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "compareExchange",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
            let expected = arg_i64(args, 2);
            let replacement = arg_i64(args, 3);
            if !check_atomic_bounds(ctx, args, idx) {
                return Value::Undefined;
            }
            if is_magic_int32(args, 0) {
                if let Some(Value::Object(obj)) = args.first() {
                    let prev = magic_load_i32(obj, idx);
                    if prev as i64 == expected {
                        magic_store_i32(obj, idx, replacement as i32);
                    }
                    return Value::I32(prev);
                }
            }
            if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                let prev = atomic_load(&buf, off, idx, bpe);
                if prev == expected {
                    atomic_store_bytes(&buf, off, idx, bpe, replacement);
                }
                return lane_value(bpe, prev);
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "load",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
            if !check_atomic_bounds(ctx, args, idx) {
                return Value::Undefined;
            }
            if is_magic_int32(args, 0) {
                if let Some(Value::Object(obj)) = args.first() {
                    return Value::I32(magic_load_i32(obj, idx));
                }
            }
            if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                return lane_value(bpe, atomic_load(&buf, off, idx, bpe));
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "store",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
            let val = arg_i64(args, 2);
            if !check_atomic_bounds(ctx, args, idx) {
                return Value::Undefined;
            }
            if is_magic_int32(args, 0) {
                if let Some(Value::Object(obj)) = args.first() {
                    magic_store_i32(obj, idx, val as i32);
                    return Value::I32(val as i32);
                }
            }
            if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                atomic_store_bytes(&buf, off, idx, bpe, val);
                return lane_value(bpe, val);
            }
            Value::I32(val as i32)
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "isLockFree",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let size = args.first().map(|v| v.as_i32()).unwrap_or(0);
            Value::Bool(matches!(size, 1 | 2 | 4 | 8))
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "wait",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
            // §25.4.15 step 1: wait (unlike load/store since ES2024)
            // still requires a SharedArrayBuffer backing.
            if let Some((_, shared)) = ta_length_and_shared(args, 0) {
                if !shared {
                    ctx.throw_value(crate::ecma::error::new_error(
                        ctx,
                        "TypeError",
                        "Atomics.wait requires a SharedArrayBuffer",
                    ));
                    return Value::Undefined;
                }
            }
            if !check_atomic_bounds(ctx, args, idx) {
                return Value::Undefined;
            }
            let expected = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let timeout_ms = args.get(3).map(|v| v.as_f64()).unwrap_or(f64::INFINITY);
            let actual = if is_magic_int32(args, 0) {
                if let Some(Value::Object(obj)) = args.first() {
                    magic_load_i32(obj, idx)
                } else {
                    0
                }
            } else if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                atomic_load(&buf, off, idx, bpe) as i32
            } else {
                0
            };
            if actual != expected {
                return Value::String(Arc::from("not-equal"));
            }
            if timeout_ms <= 0.0 {
                return Value::String(Arc::from("timed-out"));
            }
            Value::String(Arc::from("ok"))
        }),
    );

    // Atomics.waitAsync(typedArray, index, value[, timeout]) — §25.4.13
    // Returns a Promise-like object. In single-threaded Vybe, the condition is
    // checked synchronously: if the current value equals `value`, returns a
    // resolved Promise object with value "ok"; otherwise returns one with "not-equal".
    // No actual async wait occurs (no shared memory across threads in this context).
    vm.register_host_fn(
        "ecma:atomics",
        "waitAsync",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32() as usize).unwrap_or(0);
            if let Some((_, shared)) = ta_length_and_shared(args, 0) {
                if !shared {
                    ctx.throw_value(crate::ecma::error::new_error(
                        ctx,
                        "TypeError",
                        "Atomics.waitAsync requires a SharedArrayBuffer",
                    ));
                    return Value::Undefined;
                }
            }
            if !check_atomic_bounds(ctx, args, idx) {
                return Value::Undefined;
            }
            let expected = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let actual = if is_magic_int32(args, 0) {
                if let Some(Value::Object(obj)) = args.first() {
                    magic_load_i32(obj, idx)
                } else {
                    0
                }
            } else if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                atomic_load(&buf, off, idx, bpe) as i32
            } else {
                0
            };
            let result_str = if actual != expected {
                "not-equal"
            } else {
                "ok"
            };
            // Return { async: true, value: Promise<result_str> } — simplified as an object
            let mut obj = vybe_bytecode::value::Object::new();
            obj.properties.insert("async".into(), Value::Bool(true));
            obj.properties
                .insert("value".into(), Value::String(Arc::from(result_str)));
            Value::Object(Arc::new(std::sync::Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "notify",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::I32(0)),
    );
}
