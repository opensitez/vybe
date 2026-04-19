//! # `wasm:js-arraybuffer`, `wasm:js-sharedarraybuffer`, `wasm:js-dataview`
//!
//! Native Rust impls satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_arraybuffer_builtins.rs` per
//! ECMA-262 §25.1 / §25.2 / §25.3.
//!
//! ## Storage (Phase B4)
//!
//! `ObjectKind::ArrayBuffer(ArrayBufferState)` where `bytes` is an
//! `Arc<Mutex<Vec<u8>>>`. The shared `Arc` lets `DataView` and
//! every `TypedArray` view reference the **same byte storage** —
//! writes through any view are observable through every other view,
//! matching ECMA-262's buffer-sharing contract.
//!
//! Migrated from the previous "Vec of `Value::I32`-boxed bytes" MVP.
//! Memory density: 8× denser (1 byte vs 8 bytes per `Value::I32` on
//! 64-bit). Read / write speed: native byte slicing instead of a
//! two-level indirection through Values.
//!
//! DataView still uses `ObjectKind::Ordinary` with properties
//! carrying buffer / offset / length — a dedicated
//! `ObjectKind::DataView` variant lands in the next B4 sub-pass,
//! but the performance story is already driven by the shared-byte
//! backing we put in place here.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{ArrayBufferState, Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

const DV_TAG: &str = "__vybe_js_dataview";
const DV_BUFFER_PROP: &str = "__vybe_dv_buffer";
const DV_OFFSET_PROP: &str = "__vybe_dv_offset";
const DV_LENGTH_PROP: &str = "__vybe_dv_length";

// ── ArrayBuffer / SharedArrayBuffer construction ──────────────────────

fn new_arraybuffer(byte_length: i32, max_byte_length: i32, resizable: bool, shared: bool) -> Value {
    let n = byte_length.max(0) as usize;
    let max = max_byte_length.max(byte_length).max(0) as usize;
    let bytes = Arc::new(Mutex::new(vec![0u8; n]));
    let state = ArrayBufferState {
        bytes,
        max_byte_length: max,
        resizable,
        detached: false,
        shared,
    };
    let mut obj = Object::new();
    obj.kind = ObjectKind::ArrayBuffer(state);
    obj.properties.insert("byteLength".into(), Value::I32(n as i32));
    obj.properties.insert("maxByteLength".into(), Value::I32(max as i32));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_arraybuffer(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::ArrayBuffer(_)) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

/// Read-only snapshot of the buffer's current byte length. Call with
/// the object lock released.
fn ab_byte_length_of(obj: &Arc<Mutex<Object>>) -> usize {
    let o = obj.lock().unwrap();
    if let ObjectKind::ArrayBuffer(ref state) = o.kind {
        return state.bytes.lock().unwrap().len();
    }
    0
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
            new_arraybuffer(n, n, false, false)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "newResizable",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_i32()).unwrap_or(n);
            new_arraybuffer(n, max, true, false)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "byteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                return Value::I32(ab_byte_length_of(&ab) as i32);
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "maxByteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(state.max_byte_length as i32);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "resizable",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(if state.resizable { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "detached",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(if state.detached { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "slice",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    let src = state.bytes.lock().unwrap();
                    let len = src.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    let slice: Vec<u8> = if s < e { src[s..e].to_vec() } else { Vec::new() };
                    drop(src);
                    let slice_len = slice.len();
                    let new_state = ArrayBufferState {
                        bytes: Arc::new(Mutex::new(slice)),
                        max_byte_length: slice_len,
                        resizable: false,
                        detached: false,
                        shared: false,
                    };
                    let mut new_obj = Object::new();
                    new_obj.kind = ObjectKind::ArrayBuffer(new_state);
                    new_obj.properties.insert("byteLength".into(), Value::I32(slice_len as i32));
                    new_obj.properties.insert("maxByteLength".into(), Value::I32(slice_len as i32));
                    return Value::Object(Arc::new(Mutex::new(new_obj)));
                }
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "resize",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let new_len = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                let mut o = ab.lock().unwrap();
                let ObjectKind::ArrayBuffer(ref mut state) = o.kind else {
                    return Value::Null;
                };
                // Per ECMA-262 §25.1.5.3: RangeError when non-resizable
                // or exceeds maxByteLength. MVP: silent no-op.
                if !state.resizable || new_len > state.max_byte_length {
                    return Value::Null;
                }
                let mut bytes = state.bytes.lock().unwrap();
                bytes.resize(new_len, 0);
                drop(bytes);
                o.properties.insert("byteLength".into(), Value::I32(new_len as i32));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "transfer",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let requested = args.get(1).map(|v| v.as_i32()).unwrap_or(-1);
                let mut o = ab.lock().unwrap();
                let taken_bytes = if let ObjectKind::ArrayBuffer(ref mut state) = o.kind {
                    let mut src = state.bytes.lock().unwrap();
                    let taken = std::mem::take(&mut *src);
                    drop(src);
                    state.detached = true;
                    taken
                } else {
                    return Value::Null;
                };
                o.properties.insert("byteLength".into(), Value::I32(0));
                drop(o);

                let target_len = if requested < 0 { taken_bytes.len() } else { requested.max(0) as usize };
                let mut new_bytes = taken_bytes;
                new_bytes.resize(target_len, 0);
                let new_state = ArrayBufferState {
                    bytes: Arc::new(Mutex::new(new_bytes)),
                    max_byte_length: target_len,
                    resizable: false,
                    detached: false,
                    shared: false,
                };
                let mut new_obj = Object::new();
                new_obj.kind = ObjectKind::ArrayBuffer(new_state);
                new_obj.properties.insert("byteLength".into(), Value::I32(target_len as i32));
                new_obj.properties.insert("maxByteLength".into(), Value::I32(target_len as i32));
                return Value::Object(Arc::new(Mutex::new(new_obj)));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "transferToFixedLength",
        Box::new(|_ctx, args| {
            // Same as transfer() for MVP — both produce non-resizable.
            if let Some(ab) = is_arraybuffer(args, 0) {
                let requested = args.get(1).map(|v| v.as_i32()).unwrap_or(-1);
                let mut o = ab.lock().unwrap();
                let taken_bytes = if let ObjectKind::ArrayBuffer(ref mut state) = o.kind {
                    let mut src = state.bytes.lock().unwrap();
                    let taken = std::mem::take(&mut *src);
                    drop(src);
                    state.detached = true;
                    taken
                } else {
                    return Value::Null;
                };
                o.properties.insert("byteLength".into(), Value::I32(0));
                drop(o);

                let target_len = if requested < 0 { taken_bytes.len() } else { requested.max(0) as usize };
                let mut new_bytes = taken_bytes;
                new_bytes.resize(target_len, 0);
                let new_state = ArrayBufferState {
                    bytes: Arc::new(Mutex::new(new_bytes)),
                    max_byte_length: target_len,
                    resizable: false,
                    detached: false,
                    shared: false,
                };
                let mut new_obj = Object::new();
                new_obj.kind = ObjectKind::ArrayBuffer(new_state);
                new_obj.properties.insert("byteLength".into(), Value::I32(target_len as i32));
                new_obj.properties.insert("maxByteLength".into(), Value::I32(target_len as i32));
                return Value::Object(Arc::new(Mutex::new(new_obj)));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-arraybuffer", "isView",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if o.properties.get(DV_TAG).is_some() {
                    return Value::I32(1);
                }
                // Phase B4 continuation: check ObjectKind::TypedArray
                // once that variant lands.
            }
            Value::I32(0)
        }));
}

// ── SharedArrayBuffer ─────────────────────────────────────────────────

fn register_sharedarraybuffer(vm: &mut VM) {
    vm.register_host_fn("wasm:js-sharedarraybuffer", "new",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            new_arraybuffer(n, n, false, true)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "newGrowable",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_i32()).unwrap_or(n);
            new_arraybuffer(n, max, true, true)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "byteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                return Value::I32(ab_byte_length_of(&ab) as i32);
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "maxByteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(state.max_byte_length as i32);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "growable",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(if state.resizable { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "slice",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    let src = state.bytes.lock().unwrap();
                    let len = src.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    let slice: Vec<u8> = if s < e { src[s..e].to_vec() } else { Vec::new() };
                    let slice_len = slice.len();
                    drop(src);
                    let new_state = ArrayBufferState {
                        bytes: Arc::new(Mutex::new(slice)),
                        max_byte_length: slice_len,
                        resizable: false,
                        detached: false,
                        shared: true,
                    };
                    let mut new_obj = Object::new();
                    new_obj.kind = ObjectKind::ArrayBuffer(new_state);
                    new_obj.properties.insert("byteLength".into(), Value::I32(slice_len as i32));
                    new_obj.properties.insert("maxByteLength".into(), Value::I32(slice_len as i32));
                    return Value::Object(Arc::new(Mutex::new(new_obj)));
                }
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-sharedarraybuffer", "grow",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let new_len = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                let mut o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref mut state) = o.kind {
                    // Spec: only grow-in-place, never shrink.
                    if !state.resizable || new_len > state.max_byte_length {
                        return Value::Null;
                    }
                    let mut bytes = state.bytes.lock().unwrap();
                    if new_len >= bytes.len() {
                        bytes.resize(new_len, 0);
                    }
                    let new_byte_len = bytes.len() as i32;
                    drop(bytes);
                    o.properties.insert("byteLength".into(), Value::I32(new_byte_len));
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

/// Resolve the DataView's (buffer_arc, base_offset, view_length) for
/// a byte-range operation. Returns None if the view is misshapen.
fn dv_resolve(dv: &Arc<Mutex<Object>>) -> Option<(Arc<Mutex<Vec<u8>>>, usize, usize)> {
    let o = dv.lock().unwrap();
    let base_offset = match o.properties.get(DV_OFFSET_PROP) {
        Some(Value::I32(n)) => *n as usize,
        _ => 0,
    };
    let view_len = match o.properties.get(DV_LENGTH_PROP) {
        Some(Value::I32(n)) => *n as usize,
        _ => 0,
    };
    let buffer_obj = match o.properties.get(DV_BUFFER_PROP).cloned() {
        Some(Value::Object(b)) => b,
        _ => return None,
    };
    drop(o);
    let buf_o = buffer_obj.lock().unwrap();
    if let ObjectKind::ArrayBuffer(ref state) = buf_o.kind {
        Some((state.bytes.clone(), base_offset, view_len))
    } else {
        None
    }
}

fn dv_read_bytes(dv: &Arc<Mutex<Object>>, offset: i32, count: usize) -> Option<Vec<u8>> {
    let (bytes_arc, base, view_len) = dv_resolve(dv)?;
    if offset < 0 || (offset as usize + count) > view_len {
        return None;
    }
    let bytes = bytes_arc.lock().unwrap();
    let abs = base + offset as usize;
    if abs + count > bytes.len() {
        return None;
    }
    Some(bytes[abs..abs + count].to_vec())
}

fn dv_write_bytes(dv: &Arc<Mutex<Object>>, offset: i32, payload: &[u8]) -> bool {
    let Some((bytes_arc, base, view_len)) = dv_resolve(dv) else { return false; };
    if offset < 0 || (offset as usize + payload.len()) > view_len {
        return false;
    }
    let mut bytes = bytes_arc.lock().unwrap();
    let abs = base + offset as usize;
    if abs + payload.len() > bytes.len() {
        return false;
    }
    bytes[abs..abs + payload.len()].copy_from_slice(payload);
    true
}

fn register_dataview(vm: &mut VM) {
    vm.register_host_fn("wasm:js-dataview", "new",
        Box::new(|_ctx, args| {
            let buffer = args.first().cloned().unwrap_or(Value::Null);
            let byte_offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let byte_length_req = args.get(2).map(|v| v.as_i32()).unwrap_or(-1);
            let buffer_len = if let Value::Object(b) = &buffer {
                let o = b.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    state.bytes.lock().unwrap().len() as i32
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

    // Single-byte getters/setters (no endianness operand per spec)

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
