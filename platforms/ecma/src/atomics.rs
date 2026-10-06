use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

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
            ctx.throw_value(crate::error::new_error(
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
    crate::keys::with_index_key(idx, |key| match o.properties.get(key) {
        Some(Value::I32(v)) => *v,
        Some(Value::F64(v)) => *v as i32,
        _ => 0,
    })
}

fn magic_store_i32(obj: &Arc<Mutex<Object>>, idx: usize, val: i32) {
    let mut o = obj.lock().unwrap();
    crate::keys::with_index_key(idx, |key| {
        o.properties.insert(key.to_string(), Value::I32(val));
    });
}

fn magic_update_i32(
    obj: &Arc<Mutex<Object>>,
    idx: usize,
    update: impl FnOnce(i32) -> Option<i32>,
) -> i32 {
    let mut object = obj.lock().unwrap();
    crate::keys::with_index_key(idx, |key| {
        let previous = match object.properties.get(key) {
            Some(Value::I32(value)) => *value,
            Some(Value::F64(value)) => *value as i32,
            _ => 0,
        };
        if let Some(value) = update(previous) {
            object.properties.insert(key.to_string(), Value::I32(value));
        }
        previous
    })
}

fn atomic_load(buf: &Arc<Mutex<Vec<u8>>>, byte_offset: usize, idx: usize, bpe: usize) -> i64 {
    let data = buf.lock().unwrap();
    load_lane(&data, byte_offset, idx, bpe)
}

fn load_lane(data: &[u8], byte_offset: usize, idx: usize, bpe: usize) -> i64 {
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
    store_lane(&mut data, byte_offset, idx, bpe, val);
}

fn store_lane(data: &mut [u8], byte_offset: usize, idx: usize, bpe: usize, val: i64) {
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

/// Keep the backing lock across the entire read-modify-write operation.
/// The update is an internal arithmetic closure, never a guest callback.
fn atomic_update(
    buf: &Arc<Mutex<Vec<u8>>>,
    byte_offset: usize,
    idx: usize,
    bpe: usize,
    update: impl FnOnce(i64) -> Option<i64>,
) -> i64 {
    let mut data = buf.lock().unwrap();
    let previous = load_lane(&data, byte_offset, idx, bpe);
    if let Some(value) = update(previous) {
        store_lane(&mut data, byte_offset, idx, bpe, value);
    }
    previous
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
                            let prev = magic_update_i32(obj, idx, |previous| {
                                Some($op(previous as i64, val) as i32)
                            });
                            return Value::I32(prev);
                        }
                    }
                    if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                        // Sub-64 lanes wrap at lane width via the store's
                        // truncating cast; the op itself runs in i64.
                        let prev = atomic_update(&buf, off, idx, bpe, |previous| {
                            Some($op(previous, val))
                        });
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
                    let prev = magic_update_i32(obj, idx, |_| Some(val as i32));
                    return Value::I32(prev);
                }
            }
            if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                let prev = atomic_update(&buf, off, idx, bpe, |_| Some(val));
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
                    let prev = magic_update_i32(obj, idx, |previous| {
                        (previous as i64 == expected).then_some(replacement as i32)
                    });
                    return Value::I32(prev);
                }
            }
            if let Some((buf, off, bpe)) = typed_array_buffer(args, 0) {
                let prev = atomic_update(&buf, off, idx, bpe, |previous| {
                    (previous == expected).then_some(replacement)
                });
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
                    ctx.throw_value(crate::error::new_error(
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
                return crate::keys::string_value("not-equal");
            }
            if timeout_ms <= 0.0 {
                return crate::keys::string_value("timed-out");
            }
            crate::keys::string_value("ok")
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
                    ctx.throw_value(crate::error::new_error(
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
            let mut obj = vybe_runtime::value::Object::new();
            obj.properties.insert("async".into(), Value::Bool(true));
            obj.properties
                .insert("value".into(), crate::keys::string_value(result_str));
            Value::Object(Arc::new(std::sync::Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "ecma:atomics",
        "notify",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::I32(0)),
    );
}

#[cfg(test)]
mod atomic_rmw_speedup_tests {
    use super::*;

    #[test]
    fn updates_keep_offsets_lane_wrapping_and_compare_exchange_results() {
        for width in [1usize, 2, 4, 8] {
            let buffer = Arc::new(Mutex::new(vec![0xa5; 3 + width * 3]));
            atomic_store_bytes(&buffer, 3, 1, width, 17);
            assert_eq!(atomic_update(&buffer, 3, 1, width, |value| Some(value + 5)), 17);
            assert_eq!(atomic_load(&buffer, 3, 1, width), 22);
            assert_eq!(atomic_update(&buffer, 3, 1, width, |value| (value == 17).then_some(91)), 22);
            assert_eq!(atomic_load(&buffer, 3, 1, width), 22);
            assert_eq!(atomic_update(&buffer, 3, 1, width, |value| (value == 22).then_some(-1)), 22);
            assert_eq!(atomic_load(&buffer, 3, 1, width), if width == 1 { 255 } else { -1 });
            let data = buffer.lock().unwrap();
            assert!(data[..3 + width].iter().all(|byte| *byte == 0xa5));
            assert!(data[3 + width * 2..].iter().all(|byte| *byte == 0xa5));
        }
        let buffer = Arc::new(Mutex::new(vec![7; 4]));
        assert_eq!(atomic_update(&buffer, 0, 2, 4, |_| Some(9)), 0);
        assert_eq!(*buffer.lock().unwrap(), vec![7; 4]);
    }

    #[test]
    fn concurrent_increments_return_each_previous_value_exactly_once() {
        let buffer = Arc::new(Mutex::new(vec![0; 4]));
        let workers: Vec<_> = (0..6).map(|_| {
            let buffer = buffer.clone();
            std::thread::spawn(move || {
                (0..1000).map(|_| atomic_update(&buffer, 0, 0, 4, |value| Some(value + 1))).collect::<Vec<_>>()
            })
        }).collect();
        let mut previous: Vec<_> = workers.into_iter().flat_map(|worker| worker.join().unwrap()).collect();
        previous.sort_unstable();
        assert_eq!(previous, (0i64..6000).collect::<Vec<_>>());
        assert_eq!(atomic_load(&buffer, 0, 0, 4), 6000);
    }

    #[test]
    fn property_backing_keeps_numeric_conversion_and_conditional_write() {
        let mut object = Object::new();
        object.properties.insert("2".into(), Value::F64(17.75));
        let object = Arc::new(Mutex::new(object));
        assert_eq!(magic_update_i32(&object, 2, |_| None), 17);
        assert!(matches!(object.lock().unwrap().properties.get("2"), Some(Value::F64(_))));
        assert_eq!(magic_update_i32(&object, 2, |value| Some(value + 1)), 17);
        assert_eq!(magic_load_i32(&object, 2), 18);
        assert_eq!(magic_update_i32(&object, 3, |value| Some(value + 1)), 0);
        assert_eq!(magic_load_i32(&object, 3), 1);
    }

    #[test]
    #[ignore = "native backing-lock benchmark, not complete Atomics dispatch timing"]
    fn single_lock_rmw_benchmark() {
        for round in 0..6 {
            for optimized in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let buffer = Arc::new(Mutex::new(vec![0; 4]));
                let start = std::time::Instant::now();
                for _ in 0..100_000 {
                    let previous = if optimized {
                        atomic_update(&buffer, 0, 0, 4, |value| Some(value + 1))
                    } else {
                        let previous = atomic_load(&buffer, 0, 0, 4);
                        atomic_store_bytes(&buffer, 0, 0, 4, previous + 1);
                        previous
                    };
                    std::hint::black_box(previous);
                }
                eprintln!("round={round} optimized={optimized} elapsed={:?}", start.elapsed());
                assert_eq!(atomic_load(&buffer, 0, 0, 4), 100_000);
            }
        }
    }
}
