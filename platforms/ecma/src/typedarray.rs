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

use std::borrow::Cow;
use std::fmt::Write as _;
use std::sync::{Arc, Mutex, OnceLock};

/// The eleven `%<Type>Array.prototype%` singletons — ECMA-262 §23.2.
///
/// A FAMILY, so one keyed registry rather than eleven statics: the spec models
/// them the same way (each per-type prototype inherits from
/// `%TypedArray%.prototype`). Keyed by the constructor name
/// (`"Int32Array"`, …).
///
/// ⛔ Same defect these fix elsewhere: `ecma_globals` built a prototype for
/// each constructor while construction linked instances to nothing, so
/// `Object.getPrototypeOf(ta) === Int32Array.prototype` was `false` and
/// `ta instanceof Int32Array` fell through to a `__type` string. Both sides now
/// use THIS object. Primed in `lib::prime_shared_prototypes`.
static TYPEDARRAY_PROTOTYPES: OnceLock<
    Mutex<std::collections::HashMap<String, Arc<Mutex<Object>>>>,
> = OnceLock::new();

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

pub fn shared_typedarray_prototype(name: &str) -> Value {
    let map = TYPEDARRAY_PROTOTYPES.get_or_init(|| Mutex::new(std::collections::HashMap::new()));
    let mut guard = map.lock().unwrap();
    let proto = if let Some(proto) = guard.get(name) {
        proto.clone()
    } else {
        let mut obj = Object::new();
        obj.properties.reserve(3);
        obj.properties
            .insert("__proto__".into(), crate::object::shared_object_prototype());
        // §23.2.3.34 — `%TypedArray%.prototype[@@toStringTag]` is an
        // accessor returning the constructor name.
        obj.properties
            .insert("@@toStringTag".into(), crate::keys::string_value(name));
        obj.properties.insert(
            "__nonenum".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                crate::keys::string_value("@@toStringTag"),
            ]))),
        );
        let proto = vybe_runtime::heap::alloc(obj);
        guard.insert(name.to_owned(), proto.clone());
        proto
    };
    Value::Object(proto)
}

/// Every typed-array constructor name, for priming and for wiring.
pub const TYPED_ARRAY_NAMES: &[&str] = &[
    "Int8Array",
    "Uint8Array",
    "Uint8ClampedArray",
    "Int16Array",
    "Uint16Array",
    "Int32Array",
    "Uint32Array",
    "Float32Array",
    "Float64Array",
    "BigInt64Array",
    "BigUint64Array",
];
use vybe_runtime::value::{
    ArrayBufferState, Object, ObjectKind, TypedArrayState, TypedElemKind, Value,
};
use vybe_runtime::{HostContext, VM};

// ── Variant wiring ────────────────────────────────────────────────────

/// Ordered list of all 11 typed-array variants + their `wasm:js-*`
/// module names. The main `register` loop installs handlers for each.
const VARIANTS: &[(TypedElemKind, &str)] = &[
    (TypedElemKind::I8, "ecma:int8array"),
    (TypedElemKind::U8, "ecma:uint8array"),
    (TypedElemKind::U8Clamped, "ecma:uint8clamped"),
    (TypedElemKind::I16, "ecma:int16array"),
    (TypedElemKind::U16, "ecma:uint16array"),
    (TypedElemKind::I32, "ecma:int32array"),
    (TypedElemKind::U32, "ecma:uint32array"),
    (TypedElemKind::F32, "ecma:float32array"),
    (TypedElemKind::F64, "ecma:float64array"),
    (TypedElemKind::BigI64, "ecma:bigint64array"),
    (TypedElemKind::BigU64, "ecma:biguint64array"),
];

pub fn typed_array_name(elem: TypedElemKind) -> &'static str {
    match elem {
        TypedElemKind::I8 => "Int8Array",
        TypedElemKind::U8 => "Uint8Array",
        TypedElemKind::U8Clamped => "Uint8ClampedArray",
        TypedElemKind::I16 => "Int16Array",
        TypedElemKind::U16 => "Uint16Array",
        TypedElemKind::I32 => "Int32Array",
        TypedElemKind::U32 => "Uint32Array",
        TypedElemKind::F32 => "Float32Array",
        TypedElemKind::F64 => "Float64Array",
        TypedElemKind::BigI64 => "BigInt64Array",
        TypedElemKind::BigU64 => "BigUint64Array",
    }
}

pub fn zero_value(elem: TypedElemKind) -> Value {
    match elem {
        TypedElemKind::F32 | TypedElemKind::F64 => Value::F64(0.0),
        TypedElemKind::BigI64 | TypedElemKind::BigU64 => Value::bigint_i64(0),
        _ => Value::I32(0),
    }
}

fn push_typed_array_element_string(out: &mut String, value: Value) {
    match value {
        Value::String(text) => out.push_str(text.as_ref()),
        Value::BigInt(n) => {
            let _ = write!(out, "{}", n);
        }
        other => {
            let _ = write!(out, "{}", other);
        }
    }
}

#[inline]
fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
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
pub fn ta_live_length(ta: &TypedArrayState) -> usize {
    let buf = ta.buffer.lock().unwrap();
    ta_live_length_for_buffer(ta, buf.len())
}

fn ta_live_length_for_buffer(ta: &TypedArrayState, buffer_len: usize) -> usize {
    let bpe = ta.elem.bytes_per_element();
    if ta.byte_offset >= buffer_len {
        return 0;
    }
    let available_bytes = buffer_len - ta.byte_offset;
    let available_elems = available_bytes / bpe;
    ta.length.min(available_elems)
}

// ── Byte-level element access ─────────────────────────────────────────

/// A `Value`'s contents as raw BYTES — the ONE owner of that conversion.
///
/// TypedArray is this platform's type (§23.2), so this platform answers what
/// its bytes are. Before this existed the answer lived in FIVE places —
/// node/{crypto,zlib,buffer}, wasi/sockets, vybe/canvas — with THREE behaviors
/// (truncate vs clamp vs f64-cast) and NONE of them handled a TypedArray at
/// all: a Python `bytes` (a real Uint8Array) handed to zlib or a socket
/// silently decoded as EMPTY.
///
/// Conversion is §7.1.10 ToUint8 (truncate toward zero, wrap mod 2^8) — the
/// same rule a `Uint8Array` store applies, so an Array of numbers and the
/// typed array holding those numbers yield identical bytes.
pub fn bytes_from_value(v: &Value) -> Vec<u8> {
    match v {
        Value::String(s) => s.as_bytes().to_vec(),
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::TypedArray(ta) => {
                    // Byte-shaped views copy straight out of the buffer.
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let len = ta_live_length_for_buffer(ta, buf.len());
                    if bpe == 1 {
                        let start = ta.byte_offset.min(buf.len());
                        let end = (ta.byte_offset + len).min(buf.len());
                        return buf[start..end].to_vec();
                    }
                    let mut bytes = Vec::with_capacity(len);
                    for i in 0..len {
                        let abs = ta.byte_offset + i * bpe;
                        let value = read_element_from_locked_buffer(ta.elem, &buf, abs, bpe);
                        bytes.push((value.as_i32() & 0xFF) as u8);
                    }
                    bytes
                }
                ObjectKind::ArrayBuffer(ab) => ab.bytes.lock().unwrap().clone(),
                ObjectKind::Array(elems) => {
                    let mut bytes = Vec::with_capacity(elems.len());
                    for e in elems {
                        bytes.push((e.as_i32() & 0xFF) as u8);
                    }
                    bytes
                }
                _ => Vec::new(),
            }
        }
        _ => Vec::new(),
    }
}

/// A `Value`'s contents as f64s — same ownership story as
/// [`bytes_from_value`], for consumers whose elements exceed a byte
/// (palette entries, rect coordinates).
pub fn numbers_from_value(v: &Value) -> Vec<f64> {
    match v {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::TypedArray(ta) => {
                    let len = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut numbers = Vec::with_capacity(len);
                    for i in 0..len {
                        let abs = ta.byte_offset + i * bpe;
                        numbers.push(
                            read_element_from_locked_buffer(ta.elem, &buf, abs, bpe).as_f64(),
                        );
                    }
                    numbers
                }
                ObjectKind::ArrayBuffer(ab) => {
                    let bytes = ab.bytes.lock().unwrap();
                    let mut numbers = Vec::with_capacity(bytes.len());
                    for b in bytes.iter() {
                        numbers.push(*b as f64);
                    }
                    numbers
                }
                ObjectKind::Array(elems) => {
                    let mut numbers = Vec::with_capacity(elems.len());
                    for e in elems {
                        numbers.push(e.as_f64());
                    }
                    numbers
                }
                _ => Vec::new(),
            }
        }
        _ => Vec::new(),
    }
}

pub fn read_element(ta: &TypedArrayState, i: usize) -> Value {
    let bpe = ta.elem.bytes_per_element();
    let buf = ta.buffer.lock().unwrap();
    let abs = ta.byte_offset + i * bpe;
    read_element_from_locked_buffer(ta.elem, &buf, abs, bpe)
}

pub(crate) fn read_element_from_locked_buffer(
    elem: TypedElemKind,
    buf: &[u8],
    abs: usize,
    bpe: usize,
) -> Value {
    if abs + bpe > buf.len() {
        return zero_value(elem);
    }
    match elem {
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
            // Uint32 spans [0, 2^32) — beyond i32 — so surface as an F64
            // JS number (e.g. 4294967295), not a wrapped i32.
            Value::F64(u32::from_le_bytes(bytes) as f64)
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
            Value::bigint_i64(i64::from_le_bytes(bytes))
        }
        TypedElemKind::BigU64 => {
            let mut bytes = [0u8; 8];
            bytes.copy_from_slice(&buf[abs..abs + 8]);
            Value::bigint_u64(u64::from_le_bytes(bytes))
        }
    }
}

/// Coerce a caller-supplied value to the variant's element type per
/// ECMA-262 §23.2.3 and write it at index `i`. Out-of-bounds writes
/// are no-ops per spec (silent, not a trap).
pub fn write_element(ta: &TypedArrayState, i: usize, v: &Value) {
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
                let n = n.clamp(0.0, 255.0);
                let floor = n.floor();
                let frac = n - floor;
                if frac < 0.5 {
                    floor as i32
                } else if frac > 0.5 {
                    floor as i32 + 1
                } else {
                    let floor_i = floor as i32;
                    if floor_i % 2 == 0 {
                        floor_i
                    } else {
                        floor_i + 1
                    }
                }
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
            let n = v.as_f64();
            let val = if n.is_finite() {
                n.trunc().rem_euclid(4294967296.0) as u32 as i32
            } else {
                0
            };
            let bytes = val.to_le_bytes();
            buf[abs..abs + 4].copy_from_slice(&bytes);
        }
        TypedElemKind::U32 => {
            // §7.1.7 ToUint32: truncate toward zero, mod 2^32. `as_i32()`
            // saturated large numbers (4294967295 → i32::MAX); wrap instead.
            let n = v.as_f64();
            let val = if n.is_finite() {
                n.trunc().rem_euclid(4294967296.0) as u32
            } else {
                0
            };
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
            // §10.4.5 SetValueInBuffer: ToBigInt64 wrap of the value.
            let val = match v {
                Value::BigInt(n) => n.to_i64_wrapping(),
                Value::I64(n) => *n,
                other => other.as_i32() as i64,
            };
            let bytes = val.to_le_bytes();
            buf[abs..abs + 8].copy_from_slice(&bytes);
        }
        TypedElemKind::BigU64 => {
            // ToBigUint64 wrap.
            let val = match v {
                Value::BigInt(n) => n.to_u64_wrapping(),
                Value::I64(n) => *n as u64,
                other => other.as_i32() as u64,
            };
            let bytes = val.to_le_bytes();
            buf[abs..abs + 8].copy_from_slice(&bytes);
        }
    }
}

fn coerced_element_bytes(elem: TypedElemKind, v: &Value) -> ([u8; 8], usize) {
    let mut out = [0u8; 8];
    match elem {
        TypedElemKind::I8 => {
            out[0] = (v.as_i32() as i8) as u8;
            (out, 1)
        }
        TypedElemKind::U8 => {
            out[0] = (v.as_i32() & 0xFF) as u8;
            (out, 1)
        }
        TypedElemKind::U8Clamped => {
            let n = v.as_f64();
            let clamped = if n.is_nan() {
                0
            } else {
                let n = n.clamp(0.0, 255.0);
                let floor = n.floor();
                let frac = n - floor;
                if frac < 0.5 {
                    floor as i32
                } else if frac > 0.5 {
                    floor as i32 + 1
                } else {
                    let floor_i = floor as i32;
                    if floor_i % 2 == 0 {
                        floor_i
                    } else {
                        floor_i + 1
                    }
                }
            };
            out[0] = clamped as u8;
            (out, 1)
        }
        TypedElemKind::I16 => {
            out[..2].copy_from_slice(&(v.as_i32() as i16).to_le_bytes());
            (out, 2)
        }
        TypedElemKind::U16 => {
            out[..2].copy_from_slice(&((v.as_i32() & 0xFFFF) as u16).to_le_bytes());
            (out, 2)
        }
        TypedElemKind::I32 => {
            let n = v.as_f64();
            let val = if n.is_finite() {
                n.trunc().rem_euclid(4294967296.0) as u32 as i32
            } else {
                0
            };
            out[..4].copy_from_slice(&val.to_le_bytes());
            (out, 4)
        }
        TypedElemKind::U32 => {
            let n = v.as_f64();
            let val = if n.is_finite() {
                n.trunc().rem_euclid(4294967296.0) as u32
            } else {
                0
            };
            out[..4].copy_from_slice(&val.to_le_bytes());
            (out, 4)
        }
        TypedElemKind::F32 => {
            out[..4].copy_from_slice(&(v.as_f64() as f32).to_le_bytes());
            (out, 4)
        }
        TypedElemKind::F64 => {
            out.copy_from_slice(&v.as_f64().to_le_bytes());
            (out, 8)
        }
        TypedElemKind::BigI64 => {
            let val = match v {
                Value::BigInt(n) => n.to_i64_wrapping(),
                Value::I64(n) => *n,
                other => other.as_i32() as i64,
            };
            out.copy_from_slice(&val.to_le_bytes());
            (out, 8)
        }
        TypedElemKind::BigU64 => {
            let val = match v {
                Value::BigInt(n) => n.to_u64_wrapping(),
                Value::I64(n) => *n as u64,
                other => other.as_i32() as u64,
            };
            out.copy_from_slice(&val.to_le_bytes());
            (out, 8)
        }
    }
}

pub(crate) fn fill_typed_array_bytes(
    ta: &TypedArrayState,
    start: usize,
    end: usize,
    value: &Value,
) -> bool {
    if start >= end {
        return true;
    }
    let (bytes, bpe) = coerced_element_bytes(ta.elem, value);
    let mut buf = ta.buffer.lock().unwrap();
    let first = ta.byte_offset.saturating_add(start.saturating_mul(bpe));
    let last = ta.byte_offset.saturating_add(end.saturating_mul(bpe));
    if last > buf.len() {
        return false;
    }
    if bpe == 1 {
        buf[first..last].fill(bytes[0]);
        return true;
    }
    for offset in (first..last).step_by(bpe) {
        buf[offset..offset + bpe].copy_from_slice(&bytes[..bpe]);
    }
    true
}

pub(crate) fn write_array_values_to_typed_array_bytes(
    ta: &TypedArrayState,
    offset: usize,
    values: &[Value],
) -> bool {
    if values.is_empty() {
        return true;
    }
    let bpe = ta.elem.bytes_per_element();
    let mut buf = ta.buffer.lock().unwrap();
    let first = ta.byte_offset.saturating_add(offset.saturating_mul(bpe));
    let last = first.saturating_add(values.len().saturating_mul(bpe));
    if last > buf.len() {
        return false;
    }
    match ta.elem {
        TypedElemKind::I8 => {
            for (index, value) in values.iter().enumerate() {
                buf[first + index] = (value.as_i32() as i8) as u8;
            }
            return true;
        }
        TypedElemKind::U8 => {
            for (index, value) in values.iter().enumerate() {
                buf[first + index] = (value.as_i32() & 0xFF) as u8;
            }
            return true;
        }
        TypedElemKind::U8Clamped => {
            for (index, value) in values.iter().enumerate() {
                let n = value.as_f64();
                let clamped = if n.is_nan() {
                    0
                } else {
                    let n = n.clamp(0.0, 255.0);
                    let floor = n.floor();
                    let frac = n - floor;
                    if frac < 0.5 {
                        floor as i32
                    } else if frac > 0.5 {
                        floor as i32 + 1
                    } else {
                        let floor_i = floor as i32;
                        if floor_i % 2 == 0 {
                            floor_i
                        } else {
                            floor_i + 1
                        }
                    }
                };
                buf[first + index] = clamped as u8;
            }
            return true;
        }
        _ => {}
    }
    for (index, value) in values.iter().enumerate() {
        let (bytes, byte_count) = coerced_element_bytes(ta.elem, value);
        let pos = first + index * bpe;
        buf[pos..pos + byte_count].copy_from_slice(&bytes[..byte_count]);
    }
    true
}

// ── Construction helpers ──────────────────────────────────────────────

/// Allocate a fresh `length`-element typed array over a brand-new
/// ArrayBuffer. The ArrayBuffer is hidden inside the view.
pub fn new_typed_array(elem: TypedElemKind, length: usize) -> Value {
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
    ab_obj.properties.reserve(2);
    ab_obj.kind = ObjectKind::ArrayBuffer(ab_state);
    ab_obj
        .properties
        .insert("byteLength".into(), Value::I32(byte_length as i32));
    ab_obj
        .properties
        .insert("maxByteLength".into(), Value::I32(byte_length as i32));
    let buffer_obj = vybe_runtime::heap::alloc(ab_obj);

    let state = TypedArrayState {
        elem,
        buffer: bytes,
        buffer_obj: buffer_obj.clone(),
        byte_offset: 0,
        length,
    };
    let mut obj = Object::new();
    obj.properties.reserve(8);
    obj.kind = ObjectKind::TypedArray(state);
    obj.properties
        .insert("buffer".into(), Value::Object(buffer_obj.clone()));
    obj.properties
        .insert("length".into(), Value::I32(length as i32));
    obj.properties
        .insert("byteLength".into(), Value::I32(byte_length as i32));
    obj.properties.insert("byteOffset".into(), Value::I32(0));
    obj.properties
        .insert("BYTES_PER_ELEMENT".into(), Value::I32(bpe as i32));
    // §23.2.5.1 — AllocateTypedArray links the instance to its
    // `%<Type>Array.prototype%`.
    obj.properties.insert(
        "__proto__".into(),
        shared_typedarray_prototype(typed_array_name(elem)),
    );
    obj.properties.insert(
        "__type".into(),
        crate::keys::string_value(typed_array_name(elem)),
    );
    obj.properties.insert(
        "tostringtag".into(),
        crate::keys::string_value(typed_array_name(elem)),
    );
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Construct a view over an existing `ArrayBuffer`.
pub fn new_view_over_buffer(
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
    obj.properties.reserve(8);
    obj.kind = ObjectKind::TypedArray(state);
    obj.properties
        .insert("buffer".into(), Value::Object(buffer_obj.clone()));
    obj.properties
        .insert("length".into(), Value::I32(length as i32));
    obj.properties
        .insert("byteLength".into(), Value::I32((length * bpe) as i32));
    obj.properties
        .insert("byteOffset".into(), Value::I32(byte_offset as i32));
    obj.properties
        .insert("BYTES_PER_ELEMENT".into(), Value::I32(bpe as i32));
    // §23.2.5.1 — AllocateTypedArray links the instance to its
    // `%<Type>Array.prototype%`.
    obj.properties.insert(
        "__proto__".into(),
        shared_typedarray_prototype(typed_array_name(elem)),
    );
    obj.properties.insert(
        "__type".into(),
        crate::keys::string_value(typed_array_name(elem)),
    );
    obj.properties.insert(
        "tostringtag".into(),
        crate::keys::string_value(typed_array_name(elem)),
    );
    Value::Object(vybe_runtime::heap::alloc(obj))
}

pub fn apply_constructor_species(result: &Value, ctor: Value) {
    let (name, prototype) = match ctor {
        Value::Object(ctor_obj) => {
            let ctor_lock = ctor_obj.lock().unwrap();
            let name = match ctor_lock.properties.get("name") {
                Some(Value::String(name)) if !name.is_empty() => Some(name.to_string()),
                _ => None,
            };
            let prototype = match ctor_lock.properties.get("prototype") {
                Some(Value::Object(proto)) => Some(proto.clone()),
                _ => None,
            };
            (name, prototype)
        }
        _ => (None, None),
    };
    let Value::Object(result_obj) = result else {
        return;
    };
    let mut result_lock = result_obj.lock().unwrap();
    if !matches!(result_lock.kind, ObjectKind::TypedArray(_)) {
        return;
    }
    if let Some(proto) = prototype {
        result_lock
            .properties
            .insert("__proto__".into(), Value::Object(proto));
    }
    if let Some(name) = name {
        let types_obj = match result_lock.properties.get("__types") {
            Some(Value::Object(types)) => types.clone(),
            _ => {
                let types = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
                result_lock
                    .properties
                    .insert("__types".into(), Value::Object(types.clone()));
                types
            }
        };
        drop(result_lock);
        let mut types_lock = types_obj.lock().unwrap();
        if let ObjectKind::Array(ref mut types) = types_lock.kind {
            let value = crate::keys::string_value(&name);
            if !types.contains(&value) {
                types.push(value);
            }
        }
    }
}

pub fn apply_receiver_species(result: &Value, receiver: &Arc<Mutex<Object>>) {
    if let Some(ctor) = crate::object::proto_walk_get(receiver, "constructor") {
        apply_constructor_species(result, ctor);
    }
}

fn split_static_typed_array_receiver(args: &[Value]) -> (Value, &[Value]) {
    if let Some(Value::Object(obj)) = args.first() {
        let lock = obj.lock().unwrap();
        if matches!(
            lock.properties.get("__vybe_typed_array_ctor"),
            Some(Value::Bool(true))
        ) {
            return (Value::Object(obj.clone()), &args[1..]);
        }
    }
    (Value::Undefined, args)
}

fn copy_same_kind_typed_array_bytes(
    src: &TypedArrayState,
    dst: &TypedArrayState,
    dst_offset: usize,
    count: usize,
) -> bool {
    if src.elem != dst.elem {
        return false;
    }
    let bpe = src.elem.bytes_per_element();
    let byte_len = count.saturating_mul(bpe);
    let src_start = src.byte_offset;
    let dst_start = dst
        .byte_offset
        .saturating_add(dst_offset.saturating_mul(bpe));
    if Arc::ptr_eq(&src.buffer, &dst.buffer) {
        let mut buf = src.buffer.lock().unwrap();
        if src_start.saturating_add(byte_len) > buf.len()
            || dst_start.saturating_add(byte_len) > buf.len()
        {
            return false;
        }
        buf.copy_within(src_start..src_start + byte_len, dst_start);
        return true;
    }
    let src_buf = src.buffer.lock().unwrap();
    if src_start.saturating_add(byte_len) > src_buf.len() {
        return false;
    }
    let mut dst_buf = dst.buffer.lock().unwrap();
    if dst_start.saturating_add(byte_len) > dst_buf.len() {
        return false;
    }
    dst_buf[dst_start..dst_start + byte_len]
        .copy_from_slice(&src_buf[src_start..src_start + byte_len]);
    true
}

fn copy_typed_array_slice_bytes(
    src: &TypedArrayState,
    src_offset: usize,
    dst: &TypedArrayState,
    count: usize,
) -> bool {
    if src.elem != dst.elem {
        return false;
    }
    if count == 0 {
        return true;
    }
    let bpe = src.elem.bytes_per_element();
    let byte_len = count.saturating_mul(bpe);
    let src_start = src
        .byte_offset
        .saturating_add(src_offset.saturating_mul(bpe));
    let dst_start = dst.byte_offset;
    if Arc::ptr_eq(&src.buffer, &dst.buffer) {
        let mut buf = src.buffer.lock().unwrap();
        if src_start.saturating_add(byte_len) > buf.len()
            || dst_start.saturating_add(byte_len) > buf.len()
        {
            return false;
        }
        buf.copy_within(src_start..src_start + byte_len, dst_start);
        return true;
    }
    let src_buf = src.buffer.lock().unwrap();
    if src_start.saturating_add(byte_len) > src_buf.len() {
        return false;
    }
    let mut dst_buf = dst.buffer.lock().unwrap();
    if dst_start.saturating_add(byte_len) > dst_buf.len() {
        return false;
    }
    dst_buf[dst_start..dst_start + byte_len]
        .copy_from_slice(&src_buf[src_start..src_start + byte_len]);
    true
}

pub(crate) fn copy_within_typed_array_bytes(
    ta: &TypedArrayState,
    target: usize,
    start: usize,
    end: usize,
) {
    if start >= end {
        return;
    }
    let live = ta_live_length(ta);
    let max_copy = live.saturating_sub(target).min(end - start);
    if max_copy == 0 {
        return;
    }
    let bpe = ta.elem.bytes_per_element();
    let src_start = ta.byte_offset + start * bpe;
    let src_end = src_start + max_copy * bpe;
    let dst_start = ta.byte_offset + target * bpe;
    let mut buf = ta.buffer.lock().unwrap();
    if src_end <= buf.len() && dst_start + max_copy * bpe <= buf.len() {
        buf.copy_within(src_start..src_end, dst_start);
    }
}

fn copy_reversed_typed_array_bytes(
    src: &TypedArrayState,
    dst: &TypedArrayState,
    count: usize,
) -> bool {
    if src.elem != dst.elem {
        return false;
    }
    let bpe = src.elem.bytes_per_element();
    let byte_len = count.saturating_mul(bpe);
    if Arc::ptr_eq(&src.buffer, &dst.buffer) {
        let src_buf = src.buffer.lock().unwrap();
        let src_start = src.byte_offset;
        if src_start.saturating_add(byte_len) > src_buf.len() {
            return false;
        }
        let src_bytes = src_buf[src_start..src_start + byte_len].to_vec();
        drop(src_buf);
        let mut dst_buf = dst.buffer.lock().unwrap();
        let dst_start = dst.byte_offset;
        if dst_start.saturating_add(byte_len) > dst_buf.len() {
            return false;
        }
        for index in 0..count {
            let src_pos = (count - 1 - index) * bpe;
            let dst_pos = dst_start + index * bpe;
            dst_buf[dst_pos..dst_pos + bpe].copy_from_slice(&src_bytes[src_pos..src_pos + bpe]);
        }
        return true;
    }
    let src_buf = src.buffer.lock().unwrap();
    let src_start = src.byte_offset;
    if src_start.saturating_add(byte_len) > src_buf.len() {
        return false;
    }
    let mut dst_buf = dst.buffer.lock().unwrap();
    let dst_start = dst.byte_offset;
    if dst_start.saturating_add(byte_len) > dst_buf.len() {
        return false;
    }
    for index in 0..count {
        let src_pos = src_start + (count - 1 - index) * bpe;
        let dst_pos = dst_start + index * bpe;
        dst_buf[dst_pos..dst_pos + bpe].copy_from_slice(&src_buf[src_pos..src_pos + bpe]);
    }
    true
}

pub(crate) fn reverse_typed_array_bytes(ta: &TypedArrayState, count: usize) -> bool {
    let bpe = ta.elem.bytes_per_element();
    let byte_len = count.saturating_mul(bpe);
    let start = ta.byte_offset;
    let mut buf = ta.buffer.lock().unwrap();
    if start.saturating_add(byte_len) > buf.len() {
        return false;
    }
    for index in 0..(count / 2) {
        let left = start + index * bpe;
        let right = start + (count - 1 - index) * bpe;
        for byte in 0..bpe {
            buf.swap(left + byte, right + byte);
        }
    }
    true
}

enum TypedArraySearchResult {
    Ineligible,
    Found(usize),
    NotFound,
}

fn typed_array_integer_search_bytes(
    ta: &TypedArrayState,
    needle: &Value,
    from: usize,
    reverse: bool,
) -> TypedArraySearchResult {
    match ta.elem {
        TypedElemKind::I8
        | TypedElemKind::U8
        | TypedElemKind::U8Clamped
        | TypedElemKind::I16
        | TypedElemKind::U16
        | TypedElemKind::I32
        | TypedElemKind::U32 => {}
        _ => return TypedArraySearchResult::Ineligible,
    }
    let live = ta_live_length(ta);
    if from >= live {
        return TypedArraySearchResult::NotFound;
    }
    let numeric_needle = match needle {
        Value::I32(_) | Value::I64(_) => true,
        Value::F32(value) => value.is_finite(),
        Value::F64(value) => value.is_finite(),
        _ => false,
    };
    if !numeric_needle {
        return TypedArraySearchResult::NotFound;
    }
    let (needle_bytes, bpe) = coerced_element_bytes(ta.elem, needle);
    let start = ta.byte_offset;
    let byte_len = live.saturating_mul(bpe);
    let buf = ta.buffer.lock().unwrap();
    if start.saturating_add(byte_len) > buf.len() {
        return TypedArraySearchResult::Ineligible;
    }
    if reverse {
        for index in (0..=from.min(live - 1)).rev() {
            let pos = start + index * bpe;
            if buf[pos..pos + bpe] == needle_bytes[..bpe] {
                return TypedArraySearchResult::Found(index);
            }
        }
    } else {
        for index in from..live {
            let pos = start + index * bpe;
            if buf[pos..pos + bpe] == needle_bytes[..bpe] {
                return TypedArraySearchResult::Found(index);
            }
        }
    }
    TypedArraySearchResult::NotFound
}

fn try_fast_typed_array_set(
    ctx: &mut vybe_runtime::HostContext,
    args: &[Value],
    elem: TypedElemKind,
    offset: usize,
) -> Option<Value> {
    let Some(Value::Object(src_obj)) = args.get(1) else {
        return None;
    };
    let src_ta = {
        let src = src_obj.lock().unwrap();
        match &src.kind {
            ObjectKind::TypedArray(ta) if ta.elem == elem => ta.clone(),
            _ => return None,
        }
    };
    let dst_obj = is_typed_of(args, 0, elem)?;
    let dst_ta = {
        let dst = dst_obj.lock().unwrap();
        match &dst.kind {
            ObjectKind::TypedArray(ta) => ta.clone(),
            _ => return None,
        }
    };
    let count = ta_live_length(&src_ta);
    if offset.saturating_add(count) > ta_live_length(&dst_ta) {
        let err = crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
        ctx.throw_value(err);
        return Some(Value::Undefined);
    }
    if copy_same_kind_typed_array_bytes(&src_ta, &dst_ta, offset, count) {
        Some(Value::Null)
    } else {
        None
    }
}

// ── Public registration ───────────────────────────────────────────────

/// ECMA-262's **RelativeIndex** conversion, shared by every `%TypedArray%`
/// method that takes a start/end/target position.
///
/// A NEGATIVE index counts from the end (§23.2.3: *"if relativeStart < 0, let
/// k be max(len + relativeStart, 0)"*); a non-negative one is clamped to the
/// length. It is one rule, and it was open-coded correctly in `slice` and
/// `subarray` while `copyWithin` and `fill` clamped with a bare `.max(0)` —
/// so `arr.copyWithin(-2, 0, 2)` targeted index 0 and copied nothing anywhere
/// near where the caller asked (`[10,20,30,40]` came back unchanged, against
/// node's `[10,20,10,20]`).
///
/// ⛔ That bug sat GREEN in the corpus. Its test defers its assertion through
/// `__checkLater`, a `setTimeout(..., 0)`, which fired before the check could
/// observe anything — the same vacuous-pass this file's timer clamp
/// (`platforms/web/src/timers.rs::clamp_timeout`) exposed. A wrong answer that
/// nothing asserts is indistinguishable from a right one, which is why the
/// rule lives in one function now instead of four copies, two of them wrong.
/// `pub(crate)` because `value.rs`'s typed-array dispatch arms had the SAME
/// rule open-coded four more times. Two independent implementations that
/// happened to agree is not one rule — it is a divergence waiting for whoever
/// next fixes only the copy they can see, which is exactly the history above.
///
/// ⛔ NOT for `at` (§23.2.3.1): `at` returns `undefined` out of range where
/// these CLAMP. Same-looking arithmetic, different contract — do not unify it in.
pub(crate) fn relative_index(index: i32, len: i32) -> usize {
    let resolved = if index < 0 { len + index } else { index };
    resolved.max(0).min(len) as usize
}

pub fn register(vm: &mut VM) {
    for (elem, module) in VARIANTS {
        register_variant(vm, *elem, module);
    }
    // Uint8Array-only: base64 and hex encoding/decoding (ES2025).
    register_uint8_extras(vm);
}

fn ta_invoke_magic(cb: &Value, args: &[Value]) -> Option<Value> {
    let Value::Object(obj) = cb else {
        return None;
    };
    let o = obj.lock().unwrap();
    if let Some(Value::I32(n)) = o.properties.get("__map_mul") {
        let n = *n;
        drop(o);
        return Some(Value::I32(
            args.first().map(|v| v.as_i32()).unwrap_or(0) * n,
        ));
    }
    if let Some(Value::I32(n)) = o.properties.get("__pred_gt") {
        let n = *n;
        drop(o);
        return Some(Value::Bool(
            args.first().map(|v| v.as_i32()).unwrap_or(0) > n,
        ));
    }
    if o.properties.contains_key("__reduce_add") {
        drop(o);
        let a = args.first().map(|v| v.as_i32()).unwrap_or(0);
        let b = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
        return Some(Value::I32(a + b));
    }
    if o.properties.contains_key("__noop") {
        return Some(Value::Undefined);
    }
    None
}

fn typed_array_state_from_value(value: Option<&Value>) -> Option<TypedArrayState> {
    let Some(Value::Object(obj)) = value else {
        return None;
    };
    let object = obj.lock().unwrap();
    if let ObjectKind::TypedArray(ta) = &object.kind {
        Some(ta.clone())
    } else {
        None
    }
}

pub(crate) fn typed_array_values_snapshot(src: &TypedArrayState, len: usize) -> Vec<Value> {
    let bpe = src.elem.bytes_per_element();
    let buf = src.buffer.lock().unwrap();
    if bpe == 1 {
        let start = src.byte_offset.min(buf.len());
        let end = start.saturating_add(len).min(buf.len());
        let bytes = &buf[start..end];
        return match src.elem {
            TypedElemKind::I8 => {
                let mut values = Vec::with_capacity(bytes.len());
                for byte in bytes {
                    values.push(Value::I32(*byte as i8 as i32));
                }
                values
            }
            TypedElemKind::U8 | TypedElemKind::U8Clamped => {
                let mut values = Vec::with_capacity(bytes.len());
                for byte in bytes {
                    values.push(Value::I32(*byte as i32));
                }
                values
            }
            _ => Vec::new(),
        };
    }
    let mut values = Vec::with_capacity(len);
    for i in 0..len {
        let abs = src.byte_offset + i * bpe;
        values.push(read_element_from_locked_buffer(src.elem, &buf, abs, bpe));
    }
    values
}

fn write_typed_array_source_to_typed_array(
    dst: &TypedArrayState,
    offset: usize,
    src: &TypedArrayState,
) -> bool {
    let src_len = ta_live_length(src);
    if offset.saturating_add(src_len) > ta_live_length(dst) {
        return false;
    }
    if src.elem == dst.elem && copy_same_kind_typed_array_bytes(src, dst, offset, src_len) {
        return true;
    }
    let values = typed_array_values_snapshot(src, src_len);
    write_array_values_to_typed_array_bytes(dst, offset, &values)
}

fn ta_invoke_prepared(
    ctx: &mut HostContext,
    callback: &Value,
    prepared: &Option<crate::function::PreparedBoundCallback>,
    args: &[Value],
) -> Value {
    if let Some(value) = ta_invoke_magic(callback, args) {
        value
    } else if let Some(prepared) = prepared {
        crate::function::invoke_prepared_bound_callback(ctx, prepared, args)
    } else {
        ctx.invoke(callback, args)
    }
}

fn register_uint8_extras(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:uint8array",
        "toBase64",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let len = ta_live_length(ta); // locks+releases buf internally
                    let buf = ta.buffer.lock().unwrap();
                    let end = ta.byte_offset.saturating_add(len);
                    let encoded = if end <= buf.len() {
                        base64_encode(&buf[ta.byte_offset..end])
                    } else {
                        let mut bytes = Vec::with_capacity(len);
                        for i in 0..len {
                            bytes.push(buf.get(ta.byte_offset + i).copied().unwrap_or(0));
                        }
                        base64_encode(&bytes)
                    };
                    return owned_string_value(encoded);
                }
            }
            crate::keys::string_value("")
        }),
    );

    vm.register_host_fn(
        "ecma:uint8array",
        "fromBase64",
        Box::new(|_ctx, args| {
            let text = match args.first() {
                Some(Value::String(s)) => s.as_ref(),
                _ => return new_typed_array(TypedElemKind::U8, 0),
            };
            let bytes = base64_decode(text.trim());
            let ta = new_typed_array(TypedElemKind::U8, bytes.len());
            if let Value::Object(ref obj) = ta {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref t) = o.kind {
                    let mut buf = t.buffer.lock().unwrap();
                    let end = t.byte_offset.saturating_add(bytes.len());
                    if end <= buf.len() {
                        buf[t.byte_offset..end].copy_from_slice(&bytes);
                    }
                }
            }
            ta
        }),
    );

    vm.register_host_fn(
        "ecma:uint8array",
        "toHex",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let len = ta_live_length(ta); // locks+releases buf internally
                    let buf = ta.buffer.lock().unwrap();
                    let mut hex = String::with_capacity(len * 2);
                    const HEX: &[u8; 16] = b"0123456789abcdef";
                    let end = ta.byte_offset.saturating_add(len);
                    if end <= buf.len() {
                        for &b in &buf[ta.byte_offset..end] {
                            hex.push(HEX[(b >> 4) as usize] as char);
                            hex.push(HEX[(b & 0x0f) as usize] as char);
                        }
                    } else {
                        for i in 0..len {
                            let b = buf.get(ta.byte_offset + i).copied().unwrap_or(0);
                            hex.push(HEX[(b >> 4) as usize] as char);
                            hex.push(HEX[(b & 0x0f) as usize] as char);
                        }
                    }
                    return owned_string_value(hex);
                }
            }
            crate::keys::string_value("")
        }),
    );

    vm.register_host_fn(
        "ecma:uint8array",
        "fromHex",
        Box::new(|_ctx, args| {
            let text = match args.first() {
                Some(Value::String(s)) => s.as_ref(),
                _ => return new_typed_array(TypedElemKind::U8, 0),
            };
            let s = text.trim();
            let raw = s.as_bytes();
            let len = raw.len() / 2;
            let mut bytes: Vec<u8> = Vec::with_capacity(len);
            for i in 0..len {
                let hi = hex_nibble(raw[i * 2]);
                let lo = hex_nibble(raw[i * 2 + 1]);
                if let (Some(hi), Some(lo)) = (hi, lo) {
                    bytes.push((hi << 4) | lo);
                }
            }
            let ta = new_typed_array(TypedElemKind::U8, bytes.len());
            if let Value::Object(ref obj) = ta {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref t) = o.kind {
                    let mut buf = t.buffer.lock().unwrap();
                    let end = t.byte_offset.saturating_add(bytes.len());
                    if end <= buf.len() {
                        buf[t.byte_offset..end].copy_from_slice(&bytes);
                    }
                }
            }
            ta
        }),
    );
}

fn base64_encode(bytes: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((bytes.len() + 2) / 3 * 4);
    for chunk in bytes.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = chunk.get(1).copied().unwrap_or(0) as u32;
        let b2 = chunk.get(2).copied().unwrap_or(0) as u32;
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(CHARS[((n >> 18) & 63) as usize] as char);
        out.push(CHARS[((n >> 12) & 63) as usize] as char);
        if chunk.len() > 1 {
            out.push(CHARS[((n >> 6) & 63) as usize] as char);
        } else {
            out.push('=');
        }
        if chunk.len() > 2 {
            out.push(CHARS[(n & 63) as usize] as char);
        } else {
            out.push('=');
        }
    }
    out
}

fn base64_decode(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len() / 4 * 3);
    let mut chunk = [0u8; 4];
    let mut chunk_len = 0usize;
    for byte in s.bytes().filter(|&b| b != b'=') {
        chunk[chunk_len] = byte;
        chunk_len += 1;
        if chunk_len == 4 {
            decode_base64_chunk(&chunk, chunk_len, &mut out);
            chunk_len = 0;
        }
    }
    if chunk_len > 0 {
        decode_base64_chunk(&chunk, chunk_len, &mut out);
    }
    out
}

fn decode_base64_chunk(chunk: &[u8; 4], chunk_len: usize, out: &mut Vec<u8>) {
    let mut values = [0u8; 4];
    let mut value_len = 0usize;
    for &byte in &chunk[..chunk_len] {
        if let Some(value) = base64_value(byte) {
            values[value_len] = value;
            value_len += 1;
        }
    }
    if value_len >= 2 {
        out.push((values[0] << 2) | (values[1] >> 4));
    }
    if value_len >= 3 {
        out.push((values[1] << 4) | (values[2] >> 2));
    }
    if value_len >= 4 {
        out.push((values[2] << 6) | values[3]);
    }
}

#[inline]
fn base64_value(byte: u8) -> Option<u8> {
    match byte {
        b'A'..=b'Z' => Some(byte - b'A'),
        b'a'..=b'z' => Some(byte - b'a' + 26),
        b'0'..=b'9' => Some(byte - b'0' + 52),
        b'+' => Some(62),
        b'/' => Some(63),
        _ => None,
    }
}

fn register_variant(vm: &mut VM, elem: TypedElemKind, module: &'static str) {
    // ── Construction ────────────────────────────────────────────────

    // Unified constructor — dispatches on first-argument type, matching
    // ECMA-262 §23.2.4 TypedArray(argument) overload resolution:
    //   no arg / number  → new buffer of that length
    //   ArrayBuffer      → view over it (byteOffset, length optional)
    //   TypedArray       → copy-convert elements
    //   Array / iterable → fill from elements
    vm.register_host_fn(
        module,
        "new",
        Box::new(move |ctx, args| {
            match args.first() {
                None => new_typed_array(elem, 0),
                Some(Value::I32(n)) => new_typed_array(elem, (*n).max(0) as usize),
                Some(Value::F64(n)) => new_typed_array(elem, (*n as i64).max(0) as usize),
                Some(Value::I64(n)) => new_typed_array(elem, (*n).max(0) as usize),
                Some(Value::Object(src)) => {
                    let kind_tag = {
                        let o = src.lock().unwrap();
                        match &o.kind {
                            ObjectKind::ArrayBuffer(_) => 1,
                            ObjectKind::TypedArray(_) => 2,
                            ObjectKind::Array(_) => 3,
                            _ => 0,
                        }
                    };
                    match kind_tag {
                        1 => {
                            // ArrayBuffer view
                            let buf_len = {
                                let o = src.lock().unwrap();
                                if let ObjectKind::ArrayBuffer(ref s) = o.kind {
                                    s.bytes.lock().unwrap().len()
                                } else {
                                    0
                                }
                            };
                            let byte_offset =
                                args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                            let requested_len = args.get(2).map(|v| v.as_i32()).unwrap_or(-1);
                            let bpe = elem.bytes_per_element();
                            if byte_offset % bpe != 0 {
                                ctx.throw_value(crate::error::new_error(
                                    ctx,
                                    "RangeError",
                                    &format!(
                                        "{} byteOffset must be a multiple of {}",
                                        typed_array_name(elem),
                                        bpe
                                    ),
                                ));
                                return Value::Undefined;
                            }
                            let default_len = if byte_offset < buf_len {
                                (buf_len - byte_offset) / bpe
                            } else {
                                0
                            };
                            let length = if requested_len < 0 {
                                default_len
                            } else {
                                (requested_len as usize).min(default_len)
                            };
                            new_view_over_buffer(elem, src.clone(), byte_offset, length)
                        }
                        2 => {
                            // Copy from another TypedArray
                            let src_ta = {
                                let o = src.lock().unwrap();
                                match &o.kind {
                                    ObjectKind::TypedArray(ta) => ta.clone(),
                                    _ => return new_typed_array(elem, 0),
                                }
                            };
                            let live = ta_live_length(&src_ta);
                            let out = new_typed_array(elem, live);
                            if let Value::Object(ref o) = out {
                                let ol = o.lock().unwrap();
                                if let ObjectKind::TypedArray(ref t) = ol.kind {
                                    let _ = write_typed_array_source_to_typed_array(t, 0, &src_ta);
                                }
                            }
                            out
                        }
                        3 => {
                            // Fill from plain Array
                            let s = src.lock().unwrap();
                            if let ObjectKind::Array(ref elems) = s.kind {
                                let out = new_typed_array(elem, elems.len());
                                if let Value::Object(ref o) = out {
                                    let ol = o.lock().unwrap();
                                    if let ObjectKind::TypedArray(ref t) = ol.kind {
                                        write_array_values_to_typed_array_bytes(t, 0, elems);
                                    }
                                }
                                out
                            } else {
                                new_typed_array(elem, 0)
                            }
                        }
                        _ => new_typed_array(elem, 0),
                    }
                }
                _ => new_typed_array(elem, 0),
            }
        }),
    );

    vm.register_host_fn(
        module,
        "newWithLength",
        Box::new(move |_ctx, args| {
            let n = args
                .first()
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            new_typed_array(elem, n)
        }),
    );

    vm.register_host_fn(
        module,
        "newFromBuffer",
        Box::new(move |ctx, args| {
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
            if byte_offset % bpe != 0 {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    &format!(
                        "{} byteOffset must be a multiple of {}",
                        typed_array_name(elem),
                        bpe
                    ),
                ));
                return Value::Undefined;
            }
            let default_len = if byte_offset < buffer_byte_len {
                (buffer_byte_len - byte_offset) / bpe
            } else {
                0
            };
            let length = if requested_len < 0 {
                default_len
            } else {
                (requested_len as usize).min(default_len)
            };
            new_view_over_buffer(elem, buffer, byte_offset, length)
        }),
    );

    vm.register_host_fn(
        module,
        "newFromIterable",
        Box::new(move |_ctx, args| {
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref elems) = s.kind {
                    let ta_val = new_typed_array(elem, elems.len());
                    if let Value::Object(ref ta_obj) = ta_val {
                        let ta_lock = ta_obj.lock().unwrap();
                        if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                            write_array_values_to_typed_array_bytes(ta, 0, elems);
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }),
    );

    vm.register_host_fn(
        module,
        "newFromTypedArray",
        Box::new(move |_ctx, args| {
            // Copy + coerce elements from another typed array.
            if let Some(Value::Object(src)) = args.first() {
                let src_ta = {
                    let s = src.lock().unwrap();
                    match &s.kind {
                        ObjectKind::TypedArray(src_ta) => src_ta.clone(),
                        _ => return new_typed_array(elem, 0),
                    }
                };
                let live_len = ta_live_length(&src_ta);
                let ta_val = new_typed_array(elem, live_len);
                if let Value::Object(ref ta_obj) = ta_val {
                    let ta_lock = ta_obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                        let _ = write_typed_array_source_to_typed_array(ta, 0, &src_ta);
                        return ta_val.clone();
                    }
                }
                return ta_val;
            }
            new_typed_array(elem, 0)
        }),
    );

    vm.register_host_fn(
        module,
        "from",
        Box::new(move |ctx, args| {
            let (constructor_receiver, args) = split_static_typed_array_receiver(args);
            // §23.2.2.1 TypedArray.from(source[, mapFn]) — source is any
            // iterable/array-like. Generators are drained by the compiler before
            // this host call; ordinary iterables and array-like objects use the
            // shared ECMA iterator materializer.
            let source = args.first().cloned().unwrap_or(Value::Undefined);
            if matches!(source, Value::Null | Value::Undefined) {
                let err = crate::error::new_error(
                    ctx,
                    "TypeError",
                    "TypedArray.from source is not iterable",
                );
                ctx.throw_value(err);
                return Value::Undefined;
            }
            let has_map_fn = match args.get(1) {
                Some(map_fn) => !matches!(map_fn, Value::Null | Value::Undefined),
                None => false,
            };
            if !has_map_fn {
                if let Value::Object(src_obj) = &source {
                    let src_ta = {
                        let src = src_obj.lock().unwrap();
                        match &src.kind {
                            ObjectKind::TypedArray(ta) => Some(ta.clone()),
                            _ => None,
                        }
                    };
                    if let Some(src_ta) = src_ta {
                        let live = ta_live_length(&src_ta);
                        let ta_val = new_typed_array(elem, live);
                        if let Value::Object(ref ta_obj) = ta_val {
                            let ta_lock = ta_obj.lock().unwrap();
                            if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                                let _ = write_typed_array_source_to_typed_array(ta, 0, &src_ta);
                            }
                        }
                        let constructor_receiver =
                            if matches!(constructor_receiver, Value::Undefined) {
                                ctx.current_js_this()
                            } else {
                                constructor_receiver
                            };
                        apply_constructor_species(&ta_val, constructor_receiver);
                        return ta_val;
                    }
                }
            }
            let mut values =
                match crate::iterator::try_materialize_iterable_values(ctx, &source, false) {
                    Ok(values) => values,
                    Err(reason) => {
                        ctx.throw_value(reason);
                        return Value::Undefined;
                    }
                };

            // Optional mapFn (2nd arg): mapFn(value, index) per element.
            if let Some(map_fn) = args.get(1) {
                if has_map_fn {
                    let callable = matches!(
                        map_fn,
                        Value::Object(obj)
                            if matches!(
                                obj.lock().unwrap().kind,
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
                            )
                    );
                    if !callable {
                        let err = crate::error::new_error(
                            ctx,
                            "TypeError",
                            "TypedArray.from mapFn is not callable",
                        );
                        ctx.throw_value(err);
                        return Value::Undefined;
                    }
                    let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
                    let mut mapped = Vec::with_capacity(values.len());
                    for (i, v) in values.into_iter().enumerate() {
                        mapped.push(crate::function::invoke_with_explicit_this(
                            ctx,
                            map_fn,
                            this_arg.clone(),
                            &[v, Value::I32(i as i32)],
                        ));
                    }
                    values = mapped;
                }
            }
            let ta_val = new_typed_array(elem, values.len());
            if let Value::Object(ref ta_obj) = ta_val {
                let ta_lock = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                    write_array_values_to_typed_array_bytes(ta, 0, &values);
                }
            }
            let constructor_receiver = if matches!(constructor_receiver, Value::Undefined) {
                ctx.current_js_this()
            } else {
                constructor_receiver
            };
            apply_constructor_species(&ta_val, constructor_receiver);
            return ta_val;
        }),
    );

    vm.register_host_fn(
        module,
        "of",
        Box::new(move |ctx, args| {
            let (constructor_receiver, args) = split_static_typed_array_receiver(args);
            let ta_val = new_typed_array(elem, args.len());
            if let Value::Object(ref ta_obj) = ta_val {
                let ta_lock = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = ta_lock.kind {
                    write_array_values_to_typed_array_bytes(ta, 0, args);
                }
            }
            let constructor_receiver = if matches!(constructor_receiver, Value::Undefined) {
                ctx.current_js_this()
            } else {
                constructor_receiver
            };
            apply_constructor_species(&ta_val, constructor_receiver);
            ta_val
        }),
    );

    // ── Properties ──────────────────────────────────────────────────

    vm.register_host_fn(
        module,
        "buffer",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::Object(ta.buffer_obj.clone());
                }
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        module,
        "byteOffset",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::I32(ta.byte_offset as i32);
                }
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        module,
        "byteLength",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::I32((ta_live_length(ta) * ta.elem.bytes_per_element()) as i32);
                }
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        module,
        "length",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    return Value::I32(ta_live_length(ta) as i32);
                }
            }
            Value::I32(0)
        }),
    );

    // ── Element access ──────────────────────────────────────────────

    vm.register_host_fn(
        module,
        "get",
        Box::new(move |_ctx, args| {
            let Some(i) = args
                .get(1)
                .and_then(crate::keys::non_negative_integer_index)
            else {
                return Value::Undefined;
            };
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    if i >= ta_live_length(ta) {
                        return Value::Undefined;
                    }
                    return read_element(ta, i);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "at",
        Box::new(move |_ctx, args| {
            let raw_i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let i = if raw_i < 0 { live + raw_i } else { raw_i };
                    if i < 0 || i >= live {
                        return Value::Undefined;
                    }
                    return read_element(ta, i as usize);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "set",
        Box::new(move |ctx, args| {
            // Detect set(ta, source_array, offset) — array source copy.
            let is_array_source = matches!(args.get(1), Some(Value::Object(o)) if {
                let g = o.lock().unwrap();
                matches!(g.kind, ObjectKind::Array(_) | ObjectKind::TypedArray(_))
            });
            if is_array_source {
                // Delegate to setArray logic inline.
                let raw_offset = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
                if raw_offset < 0 {
                    let err = crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
                    ctx.throw_value(err);
                    return Value::Undefined;
                }
                let offset = raw_offset as usize;
                if let Some(result) = try_fast_typed_array_set(ctx, args, elem, offset) {
                    return result;
                }
                if let Some(Value::Object(src)) = args.get(1) {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(elems) = &s.kind {
                        if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                            let o = ta_obj.lock().unwrap();
                            if let ObjectKind::TypedArray(ref ta) = o.kind {
                                let live = ta_live_length(ta);
                                if offset.saturating_add(elems.len()) > live {
                                    let err = crate::error::new_error(
                                        ctx,
                                        "RangeError",
                                        "TypedArray set offset",
                                    );
                                    ctx.throw_value(err);
                                    return Value::Undefined;
                                }
                                write_array_values_to_typed_array_bytes(ta, offset, elems);
                            }
                        }
                        return Value::Null;
                    }
                }
                if let (Some(src_ta), Some(ta_obj)) = (
                    typed_array_state_from_value(args.get(1)),
                    is_typed_of(args, 0, elem),
                ) {
                    let o = ta_obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ref ta) = o.kind {
                        if !write_typed_array_source_to_typed_array(ta, offset, &src_ta) {
                            drop(o);
                            let err =
                                crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
                            ctx.throw_value(err);
                            return Value::Undefined;
                        }
                    }
                }
                return Value::Null;
            }
            // Single-element set(ta, index, value).
            let Some(i) = args
                .get(1)
                .and_then(crate::keys::non_negative_integer_index)
            else {
                return Value::Null;
            };
            let val = args.get(2).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    if i < ta_live_length(ta) {
                        write_element(ta, i, &val);
                    }
                }
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        module,
        "setArray",
        Box::new(move |ctx, args| {
            // (ta, source, offset) — coerces source elements to `elem`.
            let raw_offset = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if raw_offset < 0 {
                let err = crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
                ctx.throw_value(err);
                return Value::Undefined;
            }
            let offset = raw_offset as usize;
            if let Some(result) = try_fast_typed_array_set(ctx, args, elem, offset) {
                return result;
            }
            if let Some(Value::Object(src)) = args.get(1) {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(elems) = &s.kind {
                    if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                        let o = ta_obj.lock().unwrap();
                        if let ObjectKind::TypedArray(ref ta) = o.kind {
                            let live = ta_live_length(ta);
                            if offset.saturating_add(elems.len()) > live {
                                let err = crate::error::new_error(
                                    ctx,
                                    "RangeError",
                                    "TypedArray set offset",
                                );
                                ctx.throw_value(err);
                                return Value::Undefined;
                            }
                            write_array_values_to_typed_array_bytes(ta, offset, elems);
                        }
                    }
                    return Value::Null;
                }
            }
            if let (Some(src_ta), Some(ta_obj)) = (
                typed_array_state_from_value(args.get(1)),
                is_typed_of(args, 0, elem),
            ) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    if !write_typed_array_source_to_typed_array(ta, offset, &src_ta) {
                        drop(o);
                        let err =
                            crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
                        ctx.throw_value(err);
                        return Value::Undefined;
                    }
                }
            }
            Value::Null
        }),
    );

    // ── Mutators that don't change length ───────────────────────────

    vm.register_host_fn(
        module,
        "copyWithin",
        Box::new(move |_ctx, args| {
            let target = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let t = relative_index(target, live);
                    let s = relative_index(start, live);
                    let e = relative_index(end, live);
                    copy_within_typed_array_bytes(ta, t, s, e);
                    return args.first().cloned().unwrap_or(Value::Null);
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        module,
        "fill",
        Box::new(move |_ctx, args| {
            let val = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let s = relative_index(start, live);
                    let e = relative_index(end, live);
                    if fill_typed_array_bytes(ta, s, e, &val) {
                        return args.first().cloned().unwrap_or(Value::Null);
                    }
                    for i in s..e {
                        write_element(ta, i, &val);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        module,
        "reverse",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    if reverse_typed_array_bytes(ta, live) {
                        return args.first().cloned().unwrap_or(Value::Null);
                    }
                    let mut i = 0usize;
                    let mut j = live.saturating_sub(1);
                    while i < j {
                        let a = read_element(ta, i);
                        let b = read_element(ta, j);
                        write_element(ta, i, &b);
                        write_element(ta, j, &a);
                        i += 1;
                        j -= 1;
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        module,
        "toReversed",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let src_ta = ta.clone();
                    drop(o);
                    let ta_val = new_typed_array(elem, live);
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            if copy_reversed_typed_array_bytes(&src_ta, t, live) {
                            } else {
                                for (out_index, src_index) in (0..live).rev().enumerate() {
                                    let value = read_element(&src_ta, src_index);
                                    write_element(t, out_index, &value);
                                }
                            }
                        }
                    }
                    apply_receiver_species(&ta_val, &ta_obj);
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }),
    );

    vm.register_host_fn(
        module,
        "sort",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut values = Vec::with_capacity(live);
                    for i in 0..live {
                        let abs = ta.byte_offset + i * bpe;
                        values.push(read_element_from_locked_buffer(ta.elem, &buf, abs, bpe));
                    }
                    drop(buf);
                    values.sort_by(|a, b| {
                        a.as_f64()
                            .partial_cmp(&b.as_f64())
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                    let _ = write_array_values_to_typed_array_bytes(ta, 0, &values);
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        module,
        "toSorted",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut values = Vec::with_capacity(live);
                    for i in 0..live {
                        let abs = ta.byte_offset + i * bpe;
                        values.push(read_element_from_locked_buffer(ta.elem, &buf, abs, bpe));
                    }
                    drop(buf);
                    values.sort_by(|a, b| {
                        a.as_f64()
                            .partial_cmp(&b.as_f64())
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                    drop(o);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            let _ = write_array_values_to_typed_array_bytes(t, 0, &values);
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }),
    );

    // ── Slicing ─────────────────────────────────────────────────────

    vm.register_host_fn(
        module,
        "slice",
        Box::new(move |_ctx, args| {
            // Per spec: slice copies bytes into a new buffer (does
            // NOT share storage with the source).
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let s = relative_index(start, live);
                    let e = relative_index(end, live);
                    let count = e.saturating_sub(s);
                    let src_ta = ta.clone();
                    drop(o);
                    let ta_val = new_typed_array(elem, count);
                    if count == 0 {
                        return ta_val;
                    }
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            if !copy_typed_array_slice_bytes(&src_ta, s, t, count) {
                                let bpe = src_ta.elem.bytes_per_element();
                                let buf = src_ta.buffer.lock().unwrap();
                                let mut values = Vec::with_capacity(count);
                                for i in s..e {
                                    let abs = src_ta.byte_offset + i * bpe;
                                    values.push(read_element_from_locked_buffer(
                                        src_ta.elem,
                                        &buf,
                                        abs,
                                        bpe,
                                    ));
                                }
                                drop(buf);
                                let _ = write_array_values_to_typed_array_bytes(t, 0, &values);
                            }
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }),
    );

    vm.register_host_fn(
        module,
        "subarray",
        Box::new(move |_ctx, args| {
            // Per spec: subarray shares storage with the source.
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta) as i32;
                    let s = relative_index(start, live);
                    let e = relative_index(end, live);
                    let sub_len = if s < e { e - s } else { 0 };
                    let buffer_obj = ta.buffer_obj.clone();
                    let bpe = ta.elem.bytes_per_element();
                    let abs_offset = ta.byte_offset + s * bpe;
                    drop(o);
                    return new_view_over_buffer(elem, buffer_obj, abs_offset, sub_len);
                }
            }
            new_typed_array(elem, 0)
        }),
    );

    // ── Search ──────────────────────────────────────────────────────

    vm.register_host_fn(
        module,
        "indexOf",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            let from = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    match typed_array_integer_search_bytes(ta, &needle, from, false) {
                        TypedArraySearchResult::Found(index) => return Value::I32(index as i32),
                        TypedArraySearchResult::NotFound => return Value::I32(-1),
                        TypedArraySearchResult::Ineligible => {}
                    }
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    for i in from..live {
                        let abs = ta.byte_offset + i * bpe;
                        let value = read_element_from_locked_buffer(ta.elem, &buf, abs, bpe);
                        if Value::same_value_zero(&value, &needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    vm.register_host_fn(
        module,
        "lastIndexOf",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    if live > 0 {
                        match typed_array_integer_search_bytes(ta, &needle, live - 1, true) {
                            TypedArraySearchResult::Found(index) => {
                                return Value::I32(index as i32);
                            }
                            TypedArraySearchResult::NotFound => return Value::I32(-1),
                            TypedArraySearchResult::Ineligible => {}
                        }
                    }
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    for i in (0..live).rev() {
                        let abs = ta.byte_offset + i * bpe;
                        let value = read_element_from_locked_buffer(ta.elem, &buf, abs, bpe);
                        if Value::same_value_zero(&value, &needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    vm.register_host_fn(
        module,
        "includes",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or_else(|| zero_value(elem));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    match typed_array_integer_search_bytes(ta, &needle, 0, false) {
                        TypedArraySearchResult::Found(_) => return Value::Bool(true),
                        TypedArraySearchResult::NotFound => return Value::Bool(false),
                        TypedArraySearchResult::Ineligible => {}
                    }
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    for i in 0..live {
                        let abs = ta.byte_offset + i * bpe;
                        let value = read_element_from_locked_buffer(ta.elem, &buf, abs, bpe);
                        if Value::same_value_zero(&value, &needle) {
                            return Value::Bool(true);
                        }
                    }
                }
            }
            Value::Bool(false)
        }),
    );

    // ── join / toString ─────────────────────────────────────────────

    vm.register_host_fn(
        module,
        "join",
        Box::new(move |_ctx, args| {
            let sep = args
                .get(1)
                .map(|v| match v {
                    Value::String(text) => Cow::Borrowed(text.as_ref()),
                    other => crate::keys::value_display_cow(other),
                })
                .unwrap_or(Cow::Borrowed(","));
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut out =
                        String::with_capacity(live.saturating_mul(sep.len().saturating_add(4)));
                    for i in 0..live {
                        if i > 0 {
                            out.push_str(&sep);
                        }
                        let abs = ta.byte_offset + i * bpe;
                        push_typed_array_element_string(
                            &mut out,
                            read_element_from_locked_buffer(ta.elem, &buf, abs, bpe),
                        );
                    }
                    return owned_string_value(out);
                }
            }
            crate::keys::string_value("")
        }),
    );

    vm.register_host_fn(
        module,
        "toString",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut out = String::with_capacity(live.saturating_mul(5));
                    for i in 0..live {
                        if i > 0 {
                            out.push(',');
                        }
                        let abs = ta.byte_offset + i * bpe;
                        push_typed_array_element_string(
                            &mut out,
                            read_element_from_locked_buffer(ta.elem, &buf, abs, bpe),
                        );
                    }
                    return owned_string_value(out);
                }
            }
            crate::keys::string_value("")
        }),
    );

    vm.register_host_fn(
        module,
        "toLocaleString",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut out = String::with_capacity(live.saturating_mul(5));
                    for i in 0..live {
                        if i > 0 {
                            out.push(',');
                        }
                        let abs = ta.byte_offset + i * bpe;
                        push_typed_array_element_string(
                            &mut out,
                            read_element_from_locked_buffer(ta.elem, &buf, abs, bpe),
                        );
                    }
                    return owned_string_value(out);
                }
            }
            crate::keys::string_value("")
        }),
    );

    // ── keys / values / entries — Array snapshots ───────────────────

    vm.register_host_fn(
        module,
        "keys",
        // §23.2.3.19: returns an Array Iterator (with `.next()`), NOT a plain
        // array — `make_array_iterator` wraps the materialized snapshot.
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let mut ks = Vec::with_capacity(live);
                    for i in 0..live {
                        ks.push(Value::I32(i as i32));
                    }
                    return crate::array::make_array_iterator(ks);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );

    vm.register_host_fn(
        module,
        "values",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut vs = Vec::with_capacity(live);
                    for i in 0..live {
                        let abs = ta.byte_offset + i * bpe;
                        vs.push(read_element_from_locked_buffer(ta.elem, &buf, abs, bpe));
                    }
                    return crate::array::make_array_iterator(vs);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );

    vm.register_host_fn(
        module,
        "entries",
        Box::new(move |_ctx, args| {
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let o = ta_obj.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = o.kind {
                    let live = ta_live_length(ta);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    let mut es = Vec::with_capacity(live);
                    for i in 0..live {
                        let abs = ta.byte_offset + i * bpe;
                        es.push(crate::array::make_pair_array(
                            Value::I32(i as i32),
                            read_element_from_locked_buffer(ta.elem, &buf, abs, bpe),
                        ));
                    }
                    return crate::array::make_array_iterator(es);
                }
            }
            crate::array::make_array_iterator(Vec::new())
        }),
    );

    // ── with(i, v) ──────────────────────────────────────────────────

    vm.register_host_fn(
        module,
        "with",
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
                    let mut values = Vec::with_capacity(live as usize);
                    let bpe = ta.elem.bytes_per_element();
                    let buf = ta.buffer.lock().unwrap();
                    for k in 0..live as usize {
                        if k as i32 == idx {
                            values.push(val.clone());
                        } else {
                            let abs = ta.byte_offset + k * bpe;
                            values.push(read_element_from_locked_buffer(ta.elem, &buf, abs, bpe));
                        }
                    }
                    drop(buf);
                    drop(o);
                    let ta_val = new_typed_array(elem, values.len());
                    if let Value::Object(ref out) = ta_val {
                        let ol = out.lock().unwrap();
                        if let ObjectKind::TypedArray(ref t) = ol.kind {
                            let _ = write_array_values_to_typed_array_bytes(t, 0, &values);
                        }
                    }
                    return ta_val;
                }
            }
            new_typed_array(elem, 0)
        }),
    );

    // ── Higher-order callback methods ───────────────────────────────────

    vm.register_host_fn(
        module,
        "forEach",
        Box::new(move |ctx, args| {
            let cb = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_callback = crate::function::prepare_bound_callback(&cb);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let typed = typed_array_state_from_value(Some(&Value::Object(ta_obj)));
                let Some(ta) = typed else {
                    return Value::Undefined;
                };
                let live = ta_live_length(&ta);
                let mut callback_args = [Value::Undefined];
                for i in 0..live {
                    let v = read_element(&ta, i);
                    callback_args[0] = v;
                    if let Some(value) = ta_invoke_magic(&cb, &callback_args) {
                        let _ = value;
                    } else if let Some(prepared) = &prepared_callback {
                        let _ = crate::function::invoke_prepared_bound_callback(
                            ctx,
                            prepared,
                            &callback_args,
                        );
                    } else {
                        let _ = ctx.invoke(&cb, &callback_args);
                    }
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "map",
        Box::new(move |ctx, args| {
            let cb = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_callback = crate::function::prepare_bound_callback(&cb);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let typed = typed_array_state_from_value(Some(&Value::Object(ta_obj)));
                let Some(ta) = typed else {
                    return new_typed_array(elem, 0);
                };
                let live = ta_live_length(&ta);
                let out = new_typed_array(elem, live);
                let out_state = match &out {
                    Value::Object(out_obj) => {
                        let ol = out_obj.lock().unwrap();
                        match &ol.kind {
                            ObjectKind::TypedArray(t) => Some(t.clone()),
                            _ => None,
                        }
                    }
                    _ => None,
                };
                let mut callback_args = [Value::Undefined];
                for i in 0..live {
                    let v = read_element(&ta, i);
                    callback_args[0] = v;
                    let mapped_value = if let Some(value) = ta_invoke_magic(&cb, &callback_args) {
                        value
                    } else if let Some(prepared) = &prepared_callback {
                        crate::function::invoke_prepared_bound_callback(
                            ctx,
                            prepared,
                            &callback_args,
                        )
                    } else {
                        ctx.invoke(&cb, &callback_args)
                    };
                    if let Some(out_state) = &out_state {
                        write_element(out_state, i, &mapped_value);
                    }
                }
                return out;
            }
            new_typed_array(elem, 0)
        }),
    );

    vm.register_host_fn(
        module,
        "filter",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let typed = typed_array_state_from_value(Some(&Value::Object(ta_obj)));
                let Some(ta) = typed else {
                    return new_typed_array(elem, 0);
                };
                let live = ta_live_length(&ta);
                let mut filtered: Vec<Value> = Vec::with_capacity(live);
                let mut invoke_args = [Value::Undefined];
                for i in 0..live {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value.clone();
                    if ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        filtered.push(value);
                    }
                }
                let out = new_typed_array(elem, filtered.len());
                if let Value::Object(ref out_obj) = out {
                    let ol = out_obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ref t) = ol.kind {
                        let _ = write_array_values_to_typed_array_bytes(t, 0, &filtered);
                    }
                }
                return out;
            }
            new_typed_array(elem, 0)
        }),
    );

    vm.register_host_fn(
        module,
        "reduce",
        Box::new(move |ctx, args| {
            let reducer = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_reducer = crate::function::prepare_bound_callback(&reducer);
            let init = args.get(2).cloned();
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::Undefined;
                };
                let live = ta_live_length(&ta);
                let mut acc = match init {
                    Some(i) => i,
                    None if live > 0 => read_element(&ta, 0),
                    None => return Value::Undefined,
                };
                let start = if args.get(2).is_some() { 0 } else { 1 };
                let mut invoke_args = [Value::Undefined, Value::Undefined];
                for i in start..live {
                    let x = read_element(&ta, i);
                    invoke_args[0] = acc.clone();
                    invoke_args[1] = x;
                    acc = ta_invoke_prepared(ctx, &reducer, &prepared_reducer, &invoke_args);
                }
                return acc;
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "reduceRight",
        Box::new(move |ctx, args| {
            let reducer = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_reducer = crate::function::prepare_bound_callback(&reducer);
            let init = args.get(2).cloned();
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::Undefined;
                };
                let live = ta_live_length(&ta);
                let (mut acc, mut next_index) = match init {
                    Some(i) => (i, live.checked_sub(1)),
                    None if live > 0 => (read_element(&ta, live - 1), live.checked_sub(2)),
                    None => return Value::Undefined,
                };
                let mut invoke_args = [Value::Undefined, Value::Undefined];
                while let Some(i) = next_index {
                    let x = read_element(&ta, i);
                    invoke_args[0] = acc.clone();
                    invoke_args[1] = x;
                    acc = ta_invoke_prepared(ctx, &reducer, &prepared_reducer, &invoke_args);
                    next_index = i.checked_sub(1);
                }
                return acc;
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "some",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::Bool(false);
                };
                let live = ta_live_length(&ta);
                let mut invoke_args = [Value::Undefined];
                for i in 0..live {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value;
                    if ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        return Value::Bool(true);
                    }
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        module,
        "every",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::Bool(true);
                };
                let live = ta_live_length(&ta);
                let mut invoke_args = [Value::Undefined];
                for i in 0..live {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value;
                    if !ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        return Value::Bool(false);
                    }
                }
            }
            Value::Bool(true)
        }),
    );

    vm.register_host_fn(
        module,
        "find",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::Undefined;
                };
                let live = ta_live_length(&ta);
                let mut invoke_args = [Value::Undefined];
                for i in 0..live {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value.clone();
                    if ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        return value;
                    }
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "findIndex",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::I32(-1);
                };
                let live = ta_live_length(&ta);
                let mut invoke_args = [Value::Undefined];
                for i in 0..live {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value;
                    if ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    vm.register_host_fn(
        module,
        "findLast",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::Undefined;
                };
                let live = ta_live_length(&ta);
                let mut invoke_args = [Value::Undefined];
                for i in (0..live).rev() {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value.clone();
                    if ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        return value;
                    }
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        module,
        "findLastIndex",
        Box::new(move |ctx, args| {
            let pred = args.get(1).cloned().unwrap_or(Value::Null);
            let prepared_pred = crate::function::prepare_bound_callback(&pred);
            if let Some(ta_obj) = is_typed_of(args, 0, elem) {
                let Some(ta) = typed_array_state_from_value(Some(&Value::Object(ta_obj))) else {
                    return Value::I32(-1);
                };
                let live = ta_live_length(&ta);
                let mut invoke_args = [Value::Undefined];
                for i in (0..live).rev() {
                    let value = read_element(&ta, i);
                    invoke_args[0] = value;
                    if ta_invoke_prepared(ctx, &pred, &prepared_pred, &invoke_args).as_bool() {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }),
    );
}
