//! # `ecma:{int8,uint8,uint8clamped,int16,uint16,int32,uint32,float32,float64,bigint64,biguint64}array`
//!
//! Native Rust impls of the ECMA-262 §23.2 TypedArray family. **Not in
//! any WebAssembly CG proposal** — named `ecma:*` per the project
//! convention. The only real `wasm:js-*` names are `wasm:js-string`
//! (merged) and stage-1 `wasm:js-{number,boolean,undefined,symbol,bigint}`;
//! everything else is `ecma:*`. See `JS_BUILTIN_CONVENTIONS.md`.
//!
//! ## Storage (Phase B4 — packed-byte views)
//!
//! `ObjectKind::TypedArray { elem, buffer, buffer_obj, byte_offset,
//! length }` where `buffer` is the shared `Arc<Mutex<Vec<u8>>>` from
//! the underlying `ObjectKind::ArrayBuffer`. Every element access
//! reinterprets raw bytes at the correct offset:
//!   - `Int8Array.get(i)`  → `bytes[offset + i] as i8 as i32`
//!   - `Uint16Array.get(i)` → `i16::from_le_bytes(bytes[offset + 2i..offset + 2i + 2])`
//!   - `Float64Array.get(i)` → `f64::from_le_bytes(...)`
//!   - ... etc.
//!
//! Writes through any view mutate the shared buffer; other views
//! see the change immediately (ECMA-262 §23.2's buffer-sharing
//! contract).
//!
//! Byte order: little-endian for all multi-byte element access.
//! The spec says TypedArrays use the platform's native byte order
//! and all major JS engines + all major platforms we target are
//! little-endian, so this is consistent with v8 / SpiderMonkey.
//!
//! ## Element coercion on set
//!
//! Per ECMA-262 §23.2.3 each variant applies its own coercion:
//!   - Int8 / Int16 / Int32: truncate to bit-width (`as i8` / `as i16` / `as i32`)
//!   - Uint8 / Uint16 / Uint32: mask to bit-width
//!   - Uint8Clamped: saturating clamp to `[0, 255]` with NaN → 0
//!   - Float32 / Float64: coerce to f32 / f64 (narrowing loses precision)
//!   - BigInt64 / BigUint64: i64
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{
    ArrayBufferState, Object, ObjectKind, TypedArrayState, TypedElemKind, Value,
};
use vybe_bytecode::VM;

// ── Variant wiring ────────────────────────────────────────────────────

/// Ordered list of all 11 typed-array variants + their `wasm:js-*`
/// module names. The main `register` loop installs handlers for each.
const VARIANTS: &[(TypedElemKind, &str)] = &[
    (TypedElemKind::I8,        "ecma:int8array"),
    (TypedElemKind::U8,        "ecma:uint8array"),
    (TypedElemKind::U8Clamped, "ecma:uint8clamped"),
    (TypedElemKind::I16,       "ecma:int16array"),
    (TypedElemKind::U16,       "ecma:uint16array"),
    (TypedElemKind::I32,       "ecma:int32array"),
    (TypedElemKind::U32,       "ecma:uint32array"),
    (TypedElemKind::F32,       "ecma:float32array"),
    (TypedElemKind::F64,       "ecma:float64array"),
    (TypedElemKind::BigI64,    "ecma:bigint64array"),
    (TypedElemKind::BigU64,    "ecma:biguint64array"),
];

fn zero_value(elem: TypedElemKind) -> Value {
    match elem {
        TypedElemKind::F32 | TypedElemKind::F64 => Value::F64(0.0),
        TypedElemKind::BigI64 | TypedElemKind::BigU64 => Value::I64(0),
        _ => Value::I32(0),
    }
}

fn is_typed_of(args: &[Value], idx: usize, want: TypedElemKind) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            if ta.elem == want {
                drop(o);
                return Some(obj.clone());
            }
        }
    }
    None
}

/// Read the typed-array's current length in elements (may differ
/// from `state.length` if the underlying resizable buffer has shrunk
/// below this view's extent — per ECMA-262 §23.2.3, length then
/// reports the tracked view length, or 0 if the buffer has shrunk
/// past this view's offset).
fn ta_live_length(ta: &TypedArrayState) -> usize {
    let buf = ta.buffer.lock().unwrap();
    let bpe = ta.elem.bytes_per_element();
    if ta.byte_offset >= buf.len() { return 0; }
    let available_bytes = buf.len() - ta.byte_offset;
    let available_elems = available_bytes / bpe;
    ta.length.min(available_elems)
}

// ── Byte-level element access ─────────────────────────────────────────

fn read_element(ta: &TypedArrayState, i: usize) -> Value {
    let bpe = ta.elem.bytes_per_element();
    let buf = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset + i * bpe;
    if abs + bpe > buf.len() {
        return zero_value(ta.elem);
    }
    match ta.elem {
        TypedElemKind::I8 => Value::I32(buf[abs] as i8 as i32),
        TypedElemKind::U8 | TypedElemKind::U8Clamped => Value::I32(buf[abs] as i32),
        TypedElemKind::I16 => {
            let bytes = [buf[abs], buf[abs + 1]];
            Value::I32(i16::from_le_bytes(bytes) as i32)
        }
        TypedElemKind::U16 => {
            let bytes = [buf[abs], buf[abs + 1]];
            Value::I32(u16::from_le_bytes(bytes) as i32)
        }
        TypedElemKind::I32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            Value::I32(i32::from_le_bytes(bytes))
        }
        TypedElemKind::U32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            // u32 stored as i32 on the wire; unsigned interpretation
            // is the language's job. Matches ECMA-262 §23.2.3 which
            // treats Uint32Array elements as i32-representable only
            // via tostring/coercion at the JS side.
            Value::I32(u32::from_le_bytes(bytes) as i32)
        }
        TypedElemKind::F32 => {
            let mut bytes = [0u8; 4];
            bytes.copy_from_slice(&buf[abs..abs + 4]);
            Value::F64(f32::from_le_bytes(bytes) as f64)
        }
        TypedElemKind::F64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::F64(f64::from_le_bytes(bytes))
        }
        TypedElemKind::BigI64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::I64(i64::from_le_bytes(bytes))
        }
        TypedElemKind::BigU64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::I64(u64::from_le_bytes(bytes) as i64)
        }
    }
}

/// Coerce a caller-supplied value to the variant's element type per
/// ECMA-262 §23.2.3 and write it at index `i`. Out-of-bounds writes
/// are no-ops per spec (silent, not a trap).
fn write_element(ta: &TypedArrayState, i: usize, v: &Value) {
    let bpe = ta.elem.bytes_per_element();
    let mut buf = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset + i * bpe;
    if abs + bpe > buf.len() {
        return;
    }
    match ta.elem {
        TypedElemKind::I8 => {
            buf[abs] = (v.as_i32() as i8) as u8;
        }
        TypedElemKind::U8 => {
            buf[abs] = (v.as_i32() & 0xFF) as u8;
        }
        TypedElemKind::U8Clamped => {
            let n = v.as_f64();
            let clamped = if n.is_nan() {
                0
            } else {
                n.clamp(0.0, 255.0).round() as i32
            };
            buf[abs] = clamped as u8;
        }
        TypedElemKind::I16 => {
            let val = v.as_i32() as i16;
            let bytes = val.to_le_bytes();
            buf[abs..abs + 2].copy_from_slice(&bytes);
        }
        TypedElemKind::U16 => {
            let val = (v.as_i32() & 0xFFFF) as u16;
            let bytes = val.to_le_bytes();
            buf[abs..abs + 2].copy_from_slice(&bytes);
        }
        TypedElemKind::I32 => {
            let bytes = v.as_i32().to_le_bytes();
            buf[abs..abs + 4].copy_from_slice(&bytes);
        }
        TypedElemKind::U32 => {
            let val = v.as_i32() as u32;
            let bytes = val.to_le_bytes();
            buf[abs..abs + 4].copy_from_slice(&bytes);
        }
        TypedElemKind::F32 => {
            let val = v.as_f64() as f32;
            let bytes = val.to_le_bytes();
            buf[abs..abs + 4].copy_from_slice(&bytes);
        }
        TypedElemKind::F64 => {
            let bytes = v.as_f64().to_le_bytes();
            buf[abs..abs + 8].copy_from_slice(&bytes);
        }
        TypedElemKind::BigI64 => {
            let val = match v {
                Value::I64(n) => *n,
                other => other.as_i32() as i64,
            };
            let bytes = val.to_le_bytes();
            buf[abs..abs + 8].copy_from_slice(&bytes);
        }
        TypedElemKind::BigU64 => {
            let val = match v {
                Value::I64(n) => *n as u64,
                other => other.as_i32() as u64,
            };
            let bytes = val.to_le_bytes();
            buf[abs..abs + 8].copy_from_slice(&bytes);
        }
    }
}

// ── Construction helpers ──────────────────────────────────────────────

/// Allocate a fresh `length`-element typed array over a brand-new
/// ArrayBuffer. The ArrayBuffer is hidden inside the view.
fn new_typed_array(elem: TypedElemKind, length: usize) -> Value {
    let bpe = elem.bytes_per_element();
    let byte_length = length.saturating_mul(bpe);
    let bytes = Arc::new(Mutex::new(vec![0u8; byte_length]));

    // Construct the backing ArrayBuffer object so `.buffer` returns
    // a real externref users can pass around.
    let ab_state = ArrayBufferState {
        bytes: bytes.clone(),
        max_byte_length: byte_length,
        resizable: false,
        detached: false,
        shared: false,
    };
    let mut ab_obj = Object::new();
    ab_obj.kind = ObjectKind::ArrayBuffer(ab_state);
    ab_obj.properties.insert("byteLength".into(), Value::I32(byte_length as i32));
    ab_obj.properties.insert("maxByteLength".into(), Value::I32(byte_length as i32));
    let buffer_obj = Arc::new(Mutex::new(ab_obj));

    let state = TypedArrayState {
        elem,
        buffer: bytes,
        buffer_obj: buffer_obj.clone(),
        byte_offset: 0,
        length,
    };
    let mut obj = Object::new();
    obj.kind = ObjectKind::TypedArray(state);
    obj.properties.insert("length".into(), Value::I32(length as i32));
    obj.properties.insert("byteLength".into(), Value::I32(byte_length as i32));
    obj.properties.insert("byteOffset".into(), Value::I32(0));
    obj.properties.insert("BYTES_PER_ELEMENT".into(), Value::I32(bpe as i32));
    Value::Object(Arc::new(Mutex::new(obj)))
}

/// Construct a view over an existing `ArrayBuffer`.
fn new_view_over_buffer(
    elem: TypedElemKind,
    buffer_obj: Arc<Mutex<Object>>,
    byte_offset: usize,
    length: usize,
) -> Value {
    let bpe = elem.bytes_per_element();
    let bytes = {
        let o = buffer_obj.lock().unwrap();
        if let ObjectKind::ArrayBuffer(ref state) = o.kind {
            state.bytes.clone()
        } else {
            Arc::new(Mutex::new(Vec::new()))
        }
    };
    let state = TypedArrayState {
        elem,
        buffer: bytes,
        buffer_obj: buffer_obj.clone(),
        byte_offset,
        length,
    };
    let mut obj = Object::new();
    obj.kind = ObjectKind::TypedArray(state);
    obj.properties.insert("length".into(), Value::I32(length as i32));
    obj.properties.insert("byteLength".into(), Value::I32((length * bpe) as i32));
    obj.properties.insert("byteOffset".into(), Value::I32(byte_offset as i32));
    obj.properties.insert("BYTES_PER_ELEMENT".into(), Value::I32(bpe as i32));
    Value::Object(Arc::new(Mutex::new(obj)))
}

// ── Public registration ───────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    for (elem, module) in VARIANTS {
        register_variant(vm, *elem, module);
    }
}

fn register_variant(vm: &mut VM, elem: TypedElemKind, module: &'static str) {
    // ── Construction ────────────────────────────────────────────────

    vm.register_host_fn(module, "newWithLength",
        Box::new(move |_ctx, args| {
            let n = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            new_typed_array(elem, n)
        }));

    vm.register_host_fn(module, "newFromBuffer",
        Box::new(move |_ctx, args| {
            // (buffer, byteOffset, length) — omit signalled by -1
            let buffer = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return new_typed_array(elem, 0),
            };
            let buffer_byte_len = {
                let o = buffer.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    state.bytes.lock().unwrap().len()
                } else {
                    return new_typed_array(elem, 0);
                }
            };
            let byte_offset = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let requested_len = args.get(2).map(|v| v.as_i32()).unwrap_or(-1);
            let bpe = elem.bytes_per_element();
            let default_len = if byte_offset < buffer_byte_len {
                (buffer_byte_len - byte_offset) / bpe
            } else { 0 };
            let length = if requested_len < 0 { default_len }
                         else { (requested_len as usize).min(default_len) };
            new_view_over_buffer(elem, buffer, byte_offset, length)
        }));

    vm.register_host_fn(module, "newFromIterable",
        Box::new(move |_ctx, args| {
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref elems) = s.kind {
                    let values: Vec<Value> = elems.clone();
                    drop(s);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref ta_obj) = ta_val {
                        let ta_lock = ta_obj.lock().unwrap();
                        if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(ta, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    vm.register_host_fn(module, "newFromTypedArray",
        Box::new(move |_ctx, args| {
            // Copy + coerce elements from another typed array.
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::TypedArray(ref src_ta) = s.kind {
                    let live_len = ta_live_length(src_ta);
                    let values: Vec<Value> = (0..live_len)
                        .map(|i| read_element(src_ta, i))
                        .collect();
                    drop(s);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref ta_obj) = ta_val {
                        let ta_lock = ta_obj.lock().unwrap();
                        if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(ta, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    vm.register_host_fn(module, "from",
        Box::new(move |_ctx, args| {
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                let values: Option<Vec<Value>> = match &s.kind {
                    ObjectKind::Array(elems) => Some(elems.clone()),
                    ObjectKind::TypedArray(src_ta) => Some(
                        (0..ta_live_length(src_ta))
                            .map(|i| read_element(src_ta, i))
                            .collect()
                    ),
                    _ => None,
                };
                drop(s);
                if let Some(values) = values {
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref ta_obj) = ta_val {
                        let ta_lock = ta_obj.lock().unwrap();
                        if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(ta, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    vm.register_host_fn(module, "of",
        Box::new(move |_ctx, args| {
            let values: Vec<Value> = args.to_vec();
            let ta_val = new_typed_array(elem, values.len());
            if let Value::Object(ref ta_obj) = ta_val {
                let ta_lock = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                    for (i, v) in values.iter().enumerate() {
                        write_element(ta, i, v);
                    }
                }
            }
            ta_val
        }));

    // ── Properties ──────────────────────────────────────────────────

    vm.register_host_fn(module, "buffer",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::Object(ta.buffer_obj.clone());
                }
            }
            Value::Null
        }));

    vm.register_host_fn(module, "byteOffset",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::I32(ta.byte_offset as i32);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn(module, "byteLength",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::I32((ta_live_length(ta) * ta.elem.bytes_per_element()) as i32);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn(module, "length",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::I32(ta_live_length(ta) as i32);
                }
            }
            Value::I32(0)
        }));

    // ── Element access ──────────────────────────────────────────────

    vm.register_host_fn(module, "get",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return zero_value(elem); }
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    if (i as usize) >= ta_live_length(ta) {
                        return zero_value(elem);
                    }
                    return read_element(ta, i as usize);
                }
            }
            zero_value(elem)
        }));

    vm.register_host_fn(module, "at",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return zero_value(elem); }
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    if (i as usize) >= ta_live_length(ta) {
                        return zero_value(elem);
                    }
                    return read_element(ta, i as usize);
                }
            }
            zero_value(elem)
        }));

    vm.register_host_fn(module, "set",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return Value::Null; }
            let val = args.get(2).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    if (i as usize) < ta_live_length(ta) {
                        write_element(ta, i as usize, &val);
                    }
                }
            }
            Value::Null
        }));

    vm.register_host_fn(module, "setArray",
        Box::new(move |_ctx, args| {
            // (ta, source, offset) — coerces source elements to `elem`.
            let offset = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let source_values: Vec<Value> = match args.get(1) {
                Some(Value::Object(src)) => {
                    let s = src.lock().unwrap();
                    match &s.kind {
                        ObjectKind::Array(elems) => elems.clone(),
                        ObjectKind::TypedArray(src_ta) => (0..ta_live_length(src_ta))
                            .map(|i| read_element(src_ta, i))
                            .collect(),
                        _ => Vec::new(),
                    }
                }
                _ => Vec::new(),
            };
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    for (i, v) in source_values.iter().enumerate() {
                        let idx = offset + i;
                        if idx >= live { break; }
                        write_element(ta, idx, v);
                    }
                }
            }
            Value::Null
        }));

    // ── Mutators that don't change length ───────────────────────────

    vm.register_host_fn(module, "copyWithin",
        Box::new(move |_ctx, args| {
            let target = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let t = target.max(0).min(live) as usize;
                    let s = start.max(0).min(live) as usize;
                    let e = end.max(0).min(live) as usize;
                    // Snapshot the source window before writing so
                    // overlapping copies (memmove semantics) work.
                    let snapshot: Vec<Value> = (s..e).map(|i| read_element(ta, i)).collect();
                    let max_copy = (live as usize - t).min(snapshot.len());
                    for (i, v) in snapshot[..max_copy].iter().enumerate() {
                        write_element(ta, t + i, v);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "fill",
        Box::new(move |_ctx, args| {
            let val = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let s = start.max(0).min(live) as usize;
                    let e = end.max(0).min(live) as usize;
                    for i in s..e {
                        write_element(ta, i, &val);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "reverse",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let mut i = 0usize;
                    let mut j = live.saturating_sub(1);
                    while i < j {
                        let a = read_element(ta, i);
                        let b = read_element(ta, j);
                        write_element(ta, i, &b);
                        write_element(ta, j, &a);
                        i += 1; j -= 1;
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "toReversed",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let values: Vec<Value> = (0..live).rev().map(|i| read_element(ta, i)).collect();
                    drop(o);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(t, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    vm.register_host_fn(module, "sort",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let mut values: Vec<Value> = (0..live).map(|i| read_element(ta, i)).collect();
                    values.sort_by(|a, b| a.as_f64().partial_cmp(&b.as_f64())
                        .unwrap_or(std::cmp::Ordering::Equal));
                    for (i, v) in values.iter().enumerate() {
                        write_element(ta, i, v);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "toSorted",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let mut values: Vec<Value> = (0..live).map(|i| read_element(ta, i)).collect();
                    values.sort_by(|a, b| a.as_f64().partial_cmp(&b.as_f64())
                        .unwrap_or(std::cmp::Ordering::Equal));
                    drop(o);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(t, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    // ── Slicing ─────────────────────────────────────────────────────

    vm.register_host_fn(module, "slice",
        Box::new(move |_ctx, args| {
            // Per spec: slice copies bytes into a new buffer (does
            // NOT share storage with the source).
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let s = (if start < 0 { live + start } else { start }).max(0).min(live) as usize;
                    let e = (if end < 0 { live + end } else { end }).max(0).min(live) as usize;
                    let values: Vec<Value> = if s < e {
                        (s..e).map(|i| read_element(ta, i)).collect()
                    } else {
                        Vec::new()
                    };
                    drop(o);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(t, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    vm.register_host_fn(module, "subarray",
        Box::new(move |_ctx, args| {
            // Per spec: subarray shares storage with the source.
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let s = (if start < 0 { live + start } else { start }).max(0).min(live) as usize;
                    let e = (if end < 0 { live + end } else { end }).max(0).min(live) as usize;
                    let sub_len = if s < e { e - s } else { 0 };
                    let buffer_obj = ta.buffer_obj.clone();
                    let bpe = ta.elem.bytes_per_element();
                    let abs_offset = ta.byte_offset + s * bpe;
                    drop(o);
                    return new_view_over_buffer(elem, buffer_obj, abs_offset, sub_len);
                }
            }
            new_typed_array(elem, 0)
        }));

    // ── Search ──────────────────────────────────────────────────────

    vm.register_host_fn(module, "indexOf",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            let from = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    for i in from..live {
                        if Value::same_value_zero(&read_element(ta, i), &needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }));

    vm.register_host_fn(module, "lastIndexOf",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    for i in (0..live).rev() {
                        if Value::same_value_zero(&read_element(ta, i), &needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }));

    vm.register_host_fn(module, "includes",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    for i in 0..live {
                        if Value::same_value_zero(&read_element(ta, i), &needle) {
                            return Value::I32(1);
                        }
                    }
                }
            }
            Value::I32(0)
        }));

    // ── join / toString ─────────────────────────────────────────────

    vm.register_host_fn(module, "join",
        Box::new(move |_ctx, args| {
            let sep = args.get(1).map(|v| format!("{}", v)).unwrap_or_else(|| ",".into());
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let parts: Vec<String> = (0..live).map(|i| format!("{}", read_element(ta, i))).collect();
                    return Value::String(Arc::from(parts.join(&sep).as_str()));
                }
            }
            Value::String(Arc::from(""))
        }));

    vm.register_host_fn(module, "toString",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let parts: Vec<String> = (0..live).map(|i| format!("{}", read_element(ta, i))).collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }));

    vm.register_host_fn(module, "toLocaleString",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let parts: Vec<String> = (0..live).map(|i| format!("{}", read_element(ta, i))).collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }));

    // ── keys / values / entries — Array snapshots ───────────────────

    vm.register_host_fn(module, "keys",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let ks: Vec<Value> = (0..live as i32).map(Value::I32).collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(ks))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn(module, "values",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let vs: Vec<Value> = (0..live).map(|i| read_element(ta, i)).collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(vs))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn(module, "entries",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let es: Vec<Value> = (0..live)
                        .map(|i| Value::Object(Arc::new(Mutex::new(Object::new_array(
                            vec![Value::I32(i as i32), read_element(ta, i)]
                        )))))
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(es))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    // ── with(i, v) ──────────────────────────────────────────────────

    vm.register_host_fn(module, "with",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let idx = if i < 0 { live + i } else { i };
                    if idx < 0 || idx >= live {
                        return args.first().cloned().unwrap_or(Value::Null);
                    }
                    let values: Vec<Value> = (0..live as usize).map(|k| {
                        if k as i32 == idx { val.clone() } else { read_element(ta, k) }
                    }).collect();
                    drop(o);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            for (i, v) in values.iter().enumerate() {
                                write_element(t, i, v);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }));

    // ── Higher-order callback methods (stubs — Phase B5) ────────────

    for name in &[
        "forEach", "map", "filter", "reduce", "reduceRight",
        "some", "every", "find", "findIndex", "findLast", "findLastIndex",
    ] {
        let reg_name = name.to_string();
        let closure_name = name.to_string();
        vm.register_host_fn(module, &reg_name,
            Box::new(move |_ctx, args| {
                match closure_name.as_str() {
                    "some" | "every" => Value::I32(0),
                    "findIndex" | "findLastIndex" => Value::I32(-1),
                    "find" | "findLast" => zero_value(elem),
                    "reduce" | "reduceRight" => args.get(2).cloned().unwrap_or_else(|| zero_value(elem)),
                    _ => args.first().cloned().unwrap_or(Value::Null),
                }
            }));
    }
}
