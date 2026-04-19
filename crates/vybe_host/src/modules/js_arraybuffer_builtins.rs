//! # `wasm:js-arraybuffer`, `wasm:js-sharedarraybuffer`, `wasm:js-dataview`
//!
//! Native Rust impls satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_arraybuffer_builtins.rs` per
//! ECMA-262 §25.1 / §25.2 / §25.3.
//!
//! Storage: an `Object` with `__vybe_bytes` = `Vec<u8>` carried as a
//! property, and metadata properties tracking byteLength / resizable
//! flags. Phase B4 will upgrade to dedicated `ObjectKind::ArrayBuffer`
//! / `DataView` variants with proper byte-slice backing.
//!
//! ## Byte storage convention
//!
//! We hold bytes as `Vec<u8>` inside an `Arc<Mutex<Vec<u8>>>` wrapped
//! in a `Value::Object` whose `ObjectKind` is `Ordinary`. The magic
//! property `__vybe_bytes_handle` is an i64 index into a VM-side
//! storage pool. For MVP simplicity we instead embed the Vec directly
//! via a side-channel stored in properties — this works because our
//! value model allows arbitrary nested values.
//!
//! Alternatively we could reuse `ObjectKind::Array(Vec<Value>)` with
//! each byte stored as a `Value::I32`. That's simpler and
//! functionally equivalent for MVP — opting for it here to avoid
//! adding a new `ObjectKind` variant before Phase B4.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

const AB_TAG: &str = "__vybe_js_arraybuffer";
const SAB_TAG: &str = "__vybe_js_sharedarraybuffer";
const DV_TAG: &str = "__vybe_js_dataview";
const AB_MAX_PROP: &str = "__vybe_ab_max";
const AB_RESIZABLE_PROP: &str = "__vybe_ab_resizable";
const AB_DETACHED_PROP: &str = "__vybe_ab_detached";
const DV_BUFFER_PROP: &str = "__vybe_dv_buffer";
const DV_OFFSET_PROP: &str = "__vybe_dv_offset";
const DV_LENGTH_PROP: &str = "__vybe_dv_length";

fn new_arraybuffer(byte_length: i32, max_byte_length: i32, resizable: bool) -> Value {
    // Bytes stored as Array of I32 values (one per byte). Simple and
    // avoids introducing a new ObjectKind.
    let bytes: Vec<Value> = (0..byte_length.max(0)).map(|_| Value::I32(0)).collect();
    let mut obj = Object::new_array(bytes);
    obj.properties.insert(AB_TAG.into(), Value::I32(1));
    obj.properties.insert("byteLength".into(), Value::I32(byte_length.max(0)));
    obj.properties.insert(AB_MAX_PROP.into(), Value::I32(max_byte_length));
    obj.properties.insert(AB_RESIZABLE_PROP.into(), Value::I32(if resizable { 1 } else { 0 }));
    obj.properties.insert(AB_DETACHED_PROP.into(), Value::I32(0));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn new_sharedarraybuffer(byte_length: i32, max_byte_length: i32, growable: bool) -> Value {
    let bytes: Vec<Value> = (0..byte_length.max(0)).map(|_| Value::I32(0)).collect();
    let mut obj = Object::new_array(bytes);
    obj.properties.insert(SAB_TAG.into(), Value::I32(1));
    obj.properties.insert("byteLength".into(), Value::I32(byte_length.max(0)));
    obj.properties.insert(AB_MAX_PROP.into(), Value::I32(max_byte_length));
    obj.properties.insert(AB_RESIZABLE_PROP.into(), Value::I32(if growable { 1 } else { 0 }));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_arraybuffer(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(AB_TAG).is_some() || o.properties.get(SAB_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn ab_byte_length(obj: &Arc<Mutex<Object>>) -> i32 {
    let o = obj.lock().unwrap();
    if let ObjectKind::Array(ref v) = o.kind {
        v.len() as i32
    } else {
        0
    }
}

pub fn register(vm: &mut VM) {
    register_arraybuffer(vm);
    register_sharedarraybuffer(vm);
    register_dataview(vm);
}

// ── ArrayBuffer ───────────────────────────────────────────────────────

fn register_arraybuffer(vm: &mut VM) {
    vm.register_host_fn("wasm:js-arraybuffer", "new",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            new_arraybuffer(n, n, false)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "newResizable",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_i32()).unwrap_or(n);
            new_arraybuffer(n, max, true)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "byteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                return Value::I32(ab_byte_length(&ab));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "maxByteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                return o.properties.get(AB_MAX_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "resizable",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                return o.properties.get(AB_RESIZABLE_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "detached",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                return o.properties.get(AB_DETACHED_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "slice",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
                let o = ab.lock().unwrap();
                if let ObjectKind::Array(ref bytes) = o.kind {
                    let len = bytes.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    let new_bytes: Vec<Value> = if s < e {
                        bytes[s..e].to_vec()
                    } else {
                        Vec::new()
                    };
                    let new_len = new_bytes.len() as i32;
                    let mut out = Object::new_array(new_bytes);
                    out.properties.insert(AB_TAG.into(), Value::I32(1));
                    out.properties.insert("byteLength".into(), Value::I32(new_len));
                    out.properties.insert(AB_MAX_PROP.into(), Value::I32(new_len));
                    out.properties.insert(AB_RESIZABLE_PROP.into(), Value::I32(0));
                    out.properties.insert(AB_DETACHED_PROP.into(), Value::I32(0));
                    return Value::Object(Arc::new(Mutex::new(out)));
                }
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "resize",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let new_len = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                let mut o = ab.lock().unwrap();
                // Per spec: resize fails (throws) on non-resizable. MVP: silent no-op.
                if !matches!(o.properties.get(AB_RESIZABLE_PROP), Some(Value::I32(n)) if *n != 0) {
                    return Value::Null;
                }
                if let ObjectKind::Array(ref mut bytes) = o.kind {
                    bytes.resize(new_len, Value::I32(0));
                }
                o.properties.insert("byteLength".into(), Value::I32(new_len as i32));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "transfer",
        Box::new(|_ctx, args| {
            // transfer(ab, newByteLength_or_-1) -> new ArrayBuffer;
            // original becomes detached per spec.
            if let Some(ab) = is_arraybuffer(args, 0) {
                let mut o = ab.lock().unwrap();
                let bytes_taken: Vec<Value> = if let ObjectKind::Array(ref mut b) = o.kind {
                    std::mem::take(b)
                } else {
                    Vec::new()
                };
                o.properties.insert(AB_DETACHED_PROP.into(), Value::I32(1));
                o.properties.insert("byteLength".into(), Value::I32(0));
                let requested = args.get(1).map(|v| v.as_i32()).unwrap_or(-1);
                drop(o);
                let target_len = if requested < 0 { bytes_taken.len() as i32 } else { requested };
                let mut new_bytes = bytes_taken;
                new_bytes.resize(target_len.max(0) as usize, Value::I32(0));
                let mut new_ab = Object::new_array(new_bytes);
                new_ab.properties.insert(AB_TAG.into(), Value::I32(1));
                new_ab.properties.insert("byteLength".into(), Value::I32(target_len.max(0)));
                new_ab.properties.insert(AB_MAX_PROP.into(), Value::I32(target_len.max(0)));
                new_ab.properties.insert(AB_RESIZABLE_PROP.into(), Value::I32(0));
                new_ab.properties.insert(AB_DETACHED_PROP.into(), Value::I32(0));
                return Value::Object(Arc::new(Mutex::new(new_ab)));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "transferToFixedLength",
        Box::new(|_ctx, args| {
            // Same as transfer but always produces non-resizable.
            // Our `transfer` already does that for MVP — alias.
            if let Some(ab) = is_arraybuffer(args, 0) {
                let mut o = ab.lock().unwrap();
                let bytes_taken: Vec<Value> = if let ObjectKind::Array(ref mut b) = o.kind {
                    std::mem::take(b)
                } else {
                    Vec::new()
                };
                o.properties.insert(AB_DETACHED_PROP.into(), Value::I32(1));
                o.properties.insert("byteLength".into(), Value::I32(0));
                let requested = args.get(1).map(|v| v.as_i32()).unwrap_or(-1);
                drop(o);
                let target_len = if requested < 0 { bytes_taken.len() as i32 } else { requested };
                let mut new_bytes = bytes_taken;
                new_bytes.resize(target_len.max(0) as usize, Value::I32(0));
                let mut new_ab = Object::new_array(new_bytes);
                new_ab.properties.insert(AB_TAG.into(), Value::I32(1));
                new_ab.properties.insert("byteLength".into(), Value::I32(target_len.max(0)));
                new_ab.properties.insert(AB_MAX_PROP.into(), Value::I32(target_len.max(0)));
                new_ab.properties.insert(AB_RESIZABLE_PROP.into(), Value::I32(0));
                new_ab.properties.insert(AB_DETACHED_PROP.into(), Value::I32(0));
                return Value::Object(Arc::new(Mutex::new(new_ab)));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "isView",
        Box::new(|_ctx, args| {
            // True for DataView or any typed-array view.
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if o.properties.get(DV_TAG).is_some() {
                    return Value::I32(1);
                }
                // Typed arrays will be tagged `__vybe_js_typedarray_*`;
                // MVP returns 0 for them and Phase B10 handlers add
                // the check once typed-array tagging lands.
            }
            Value::I32(0)
        }));
}

// ── SharedArrayBuffer ─────────────────────────────────────────────────

fn register_sharedarraybuffer(vm: &mut VM) {
    vm.register_host_fn("wasm:js-sharedarraybuffer", "new",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            new_sharedarraybuffer(n, n, false)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "newGrowable",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_i32()).unwrap_or(n);
            new_sharedarraybuffer(n, max, true)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "byteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                return Value::I32(ab_byte_length(&ab));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "maxByteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                return o.properties.get(AB_MAX_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "growable",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                return o.properties.get(AB_RESIZABLE_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "slice",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
                let o = ab.lock().unwrap();
                if let ObjectKind::Array(ref bytes) = o.kind {
                    let len = bytes.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    let new_bytes: Vec<Value> = if s < e {
                        bytes[s..e].to_vec()
                    } else {
                        Vec::new()
                    };
                    let new_len = new_bytes.len() as i32;
                    let mut out = Object::new_array(new_bytes);
                    out.properties.insert(SAB_TAG.into(), Value::I32(1));
                    out.properties.insert("byteLength".into(), Value::I32(new_len));
                    out.properties.insert(AB_MAX_PROP.into(), Value::I32(new_len));
                    out.properties.insert(AB_RESIZABLE_PROP.into(), Value::I32(0));
                    return Value::Object(Arc::new(Mutex::new(out)));
                }
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "grow",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let new_len = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                let mut o = ab.lock().unwrap();
                if !matches!(o.properties.get(AB_RESIZABLE_PROP), Some(Value::I32(n)) if *n != 0) {
                    return Value::Null;
                }
                if let ObjectKind::Array(ref mut bytes) = o.kind {
                    // `grow` is spec'd to only allow growth, not shrink.
                    if new_len >= bytes.len() {
                        bytes.resize(new_len, Value::I32(0));
                        o.properties.insert("byteLength".into(), Value::I32(new_len as i32));
                    }
                }
            }
            Value::Null
        }));
}

// ── DataView ──────────────────────────────────────────────────────────

fn new_dataview(buffer: Value, byte_offset: i32, byte_length: i32) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(DV_TAG.into(), Value::I32(1));
    obj.properties.insert(DV_BUFFER_PROP.into(), buffer);
    obj.properties.insert(DV_OFFSET_PROP.into(), Value::I32(byte_offset.max(0)));
    obj.properties.insert(DV_LENGTH_PROP.into(), Value::I32(byte_length.max(0)));
    obj.properties.insert("byteOffset".into(), Value::I32(byte_offset.max(0)));
    obj.properties.insert("byteLength".into(), Value::I32(byte_length.max(0)));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_dataview(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(DV_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

/// Copy bytes from the DataView's backing ArrayBuffer into a Vec<u8>
/// for the given offset+count. Returns None on out-of-bounds.
fn dv_read_bytes(dv: &Arc<Mutex<Object>>, offset: i32, count: usize) -> Option<Vec<u8>> {
    let o = dv.lock().unwrap();
    let base_offset = match o.properties.get(DV_OFFSET_PROP) {
        Some(Value::I32(n)) => *n,
        _ => 0,
    };
    let view_len = match o.properties.get(DV_LENGTH_PROP) {
        Some(Value::I32(n)) => *n,
        _ => 0,
    };
    if offset < 0 || (offset as usize + count) > view_len as usize {
        return None;
    }
    let buffer = match o.properties.get(DV_BUFFER_PROP).cloned() {
        Some(Value::Object(b)) => b,
        _ => return None,
    };
    drop(o);
    let buf = buffer.lock().unwrap();
    if let ObjectKind::Array(ref bytes) = buf.kind {
        let abs_offset = (base_offset + offset) as usize;
        let mut out = Vec::with_capacity(count);
        for i in abs_offset..abs_offset + count {
            if let Some(Value::I32(b)) = bytes.get(i) {
                out.push(*b as u8);
            } else {
                return None;
            }
        }
        Some(out)
    } else {
        None
    }
}

fn dv_write_bytes(dv: &Arc<Mutex<Object>>, offset: i32, bytes: &[u8]) -> bool {
    let o = dv.lock().unwrap();
    let base_offset = match o.properties.get(DV_OFFSET_PROP) {
        Some(Value::I32(n)) => *n,
        _ => 0,
    };
    let view_len = match o.properties.get(DV_LENGTH_PROP) {
        Some(Value::I32(n)) => *n,
        _ => 0,
    };
    if offset < 0 || (offset as usize + bytes.len()) > view_len as usize {
        return false;
    }
    let buffer = match o.properties.get(DV_BUFFER_PROP).cloned() {
        Some(Value::Object(b)) => b,
        _ => return false,
    };
    drop(o);
    let mut buf = buffer.lock().unwrap();
    if let ObjectKind::Array(ref mut arr) = buf.kind {
        let abs_offset = (base_offset + offset) as usize;
        for (i, &b) in bytes.iter().enumerate() {
            let idx = abs_offset + i;
            if idx >= arr.len() {
                return false;
            }
            arr[idx] = Value::I32(b as i32);
        }
        true
    } else {
        false
    }
}

fn register_dataview(vm: &mut VM) {
    vm.register_host_fn("wasm:js-dataview", "new",
        Box::new(|_ctx, args| {
            let buffer = args.first().cloned().unwrap_or(Value::Null);
            let byte_offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let byte_length_req = args.get(2).map(|v| v.as_i32()).unwrap_or(-1);
            // Compute default byteLength if omitted (= buffer.byteLength - offset)
            let buffer_len = if let Value::Object(b) = &buffer {
                let o = b.lock().unwrap();
                if let ObjectKind::Array(ref bytes) = o.kind {
                    bytes.len() as i32
                } else {
                    0
                }
            } else {
                0
            };
            let byte_length = if byte_length_req < 0 {
                (buffer_len - byte_offset).max(0)
            } else {
                byte_length_req
            };
            new_dataview(buffer, byte_offset, byte_length)
        }));

    vm.register_host_fn("wasm:js-dataview", "buffer",
        Box::new(|_ctx, args| {
            if let Some(dv) = is_dataview(args, 0) {
                let o = dv.lock().unwrap();
                return o.properties.get(DV_BUFFER_PROP).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-dataview", "byteOffset",
        Box::new(|_ctx, args| {
            if let Some(dv) = is_dataview(args, 0) {
                let o = dv.lock().unwrap();
                return o.properties.get(DV_OFFSET_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-dataview", "byteLength",
        Box::new(|_ctx, args| {
            if let Some(dv) = is_dataview(args, 0) {
                let o = dv.lock().unwrap();
                return o.properties.get(DV_LENGTH_PROP).cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    // Getters
    vm.register_host_fn("wasm:js-dataview", "getInt8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                if let Some(bytes) = dv_read_bytes(&dv, offset, 1) {
                    return Value::I32(bytes[0] as i8 as i32);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-dataview", "getUint8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                if let Some(bytes) = dv_read_bytes(&dv, offset, 1) {
                    return Value::I32(bytes[0] as i32);
                }
            }
            Value::I32(0)
        }));

    macro_rules! getter_multibyte {
        ($name:literal, $count:expr, $ty:ty, $le:path, $be:path, $wrap:expr) => {
            vm.register_host_fn("wasm:js-dataview", $name,
                Box::new(|_ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    let little_endian = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
                    if let Some(dv) = is_dataview(args, 0) {
                        if let Some(bytes) = dv_read_bytes(&dv, offset, $count) {
                            let mut arr = [0u8; $count];
                            arr.copy_from_slice(&bytes);
                            let val: $ty = if little_endian { $le(arr) } else { $be(arr) };
                            return $wrap(val);
                        }
                    }
                    $wrap(<$ty>::default())
                }));
        };
    }

    getter_multibyte!("getInt16", 2, i16, i16::from_le_bytes, i16::from_be_bytes, |v| Value::I32(v as i32));
    getter_multibyte!("getUint16", 2, u16, u16::from_le_bytes, u16::from_be_bytes, |v| Value::I32(v as i32));
    getter_multibyte!("getInt32", 4, i32, i32::from_le_bytes, i32::from_be_bytes, |v| Value::I32(v));
    getter_multibyte!("getUint32", 4, u32, u32::from_le_bytes, u32::from_be_bytes, |v| Value::I32(v as i32));
    getter_multibyte!("getBigInt64", 8, i64, i64::from_le_bytes, i64::from_be_bytes, |v| Value::I64(v));
    getter_multibyte!("getBigUint64", 8, u64, u64::from_le_bytes, u64::from_be_bytes, |v| Value::I64(v as i64));
    getter_multibyte!("getFloat32", 4, f32, f32::from_le_bytes, f32::from_be_bytes, |v| Value::F64(v as f64));
    getter_multibyte!("getFloat64", 8, f64, f64::from_le_bytes, f64::from_be_bytes, |v| Value::F64(v));

    // Setters
    vm.register_host_fn("wasm:js-dataview", "setInt8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                dv_write_bytes(&dv, offset, &[(val as i8) as u8]);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-dataview", "setUint8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                dv_write_bytes(&dv, offset, &[val as u8]);
            }
            Value::Null
        }));

    macro_rules! setter_multibyte {
        ($name:literal, $count:expr, $val_extract:expr, $ty:ty, $le:ident, $be:ident) => {
            vm.register_host_fn("wasm:js-dataview", $name,
                Box::new(|_ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    let val: $ty = $val_extract(args.get(2));
                    let little_endian = args.get(3).map(|v| v.as_i32()).unwrap_or(0) != 0;
                    let bytes = if little_endian { val.$le() } else { val.$be() };
                    if let Some(dv) = is_dataview(args, 0) {
                        dv_write_bytes(&dv, offset, &bytes);
                    }
                    Value::Null
                }));
        };
    }

    setter_multibyte!("setInt16", 2,
        |v: Option<&Value>| v.map(|x| x.as_i32() as i16).unwrap_or(0),
        i16, to_le_bytes, to_be_bytes);
    setter_multibyte!("setUint16", 2,
        |v: Option<&Value>| v.map(|x| x.as_i32() as u16).unwrap_or(0),
        u16, to_le_bytes, to_be_bytes);
    setter_multibyte!("setInt32", 4,
        |v: Option<&Value>| v.map(|x| x.as_i32()).unwrap_or(0),
        i32, to_le_bytes, to_be_bytes);
    setter_multibyte!("setUint32", 4,
        |v: Option<&Value>| v.map(|x| x.as_i32() as u32).unwrap_or(0),
        u32, to_le_bytes, to_be_bytes);
    setter_multibyte!("setBigInt64", 8,
        |v: Option<&Value>| v.map(|x| match x { Value::I64(n) => *n, other => other.as_i32() as i64 }).unwrap_or(0),
        i64, to_le_bytes, to_be_bytes);
    setter_multibyte!("setBigUint64", 8,
        |v: Option<&Value>| v.map(|x| match x { Value::I64(n) => *n as u64, other => other.as_i32() as u64 }).unwrap_or(0),
        u64, to_le_bytes, to_be_bytes);
    setter_multibyte!("setFloat32", 4,
        |v: Option<&Value>| v.map(|x| x.as_f64() as f32).unwrap_or(0.0),
        f32, to_le_bytes, to_be_bytes);
    setter_multibyte!("setFloat64", 8,
        |v: Option<&Value>| v.map(|x| x.as_f64()).unwrap_or(0.0),
        f64, to_le_bytes, to_be_bytes);
}
