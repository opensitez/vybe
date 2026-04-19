//! # `wasm:js-{int8,uint8,int16,…}array` host handlers
//!
//! Native Rust impls satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_typedarray_builtins.rs` per
//! ECMA-262 §23.2.
//!
//! Eleven variants share one method surface; we generate each variant's
//! handlers via a helper that closes over the variant's element type.
//!
//! Storage (MVP): an Object with `ObjectKind::Array(Vec<Value>)` of
//! boxed primitive values (`Value::I32` for integer variants, `Value::F64`
//! for float variants, `Value::I64` for BigInt variants). Phase B4 will
//! swap in dedicated `ObjectKind::TypedArray` variants backed by
//! packed slices for memory density and SIMD acceleration.
//!
//! Sign-extension / zero-extension / clamping is applied per variant
//! at the get/set boundary to match ECMA-262 semantics.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

/// Element kind discriminating signed/unsigned/clamped/float behavior.
#[derive(Copy, Clone, Debug)]
enum Elem {
    I8, U8, U8Clamped,
    I16, U16,
    I32, U32,
    F32, F64,
    BigI64, BigU64,
}

impl Elem {
    fn module(self) -> &'static str {
        match self {
            Elem::I8 => "wasm:js-int8array",
            Elem::U8 => "wasm:js-uint8array",
            Elem::U8Clamped => "wasm:js-uint8clamped",
            Elem::I16 => "wasm:js-int16array",
            Elem::U16 => "wasm:js-uint16array",
            Elem::I32 => "wasm:js-int32array",
            Elem::U32 => "wasm:js-uint32array",
            Elem::F32 => "wasm:js-float32array",
            Elem::F64 => "wasm:js-float64array",
            Elem::BigI64 => "wasm:js-bigint64array",
            Elem::BigU64 => "wasm:js-biguint64array",
        }
    }

    fn bytes_per_element(self) -> i32 {
        match self {
            Elem::I8 | Elem::U8 | Elem::U8Clamped => 1,
            Elem::I16 | Elem::U16 => 2,
            Elem::I32 | Elem::U32 | Elem::F32 => 4,
            Elem::F64 | Elem::BigI64 | Elem::BigU64 => 8,
        }
    }

    fn tag(self) -> &'static str {
        match self {
            Elem::I8 => "__vybe_ta_i8",
            Elem::U8 => "__vybe_ta_u8",
            Elem::U8Clamped => "__vybe_ta_u8c",
            Elem::I16 => "__vybe_ta_i16",
            Elem::U16 => "__vybe_ta_u16",
            Elem::I32 => "__vybe_ta_i32",
            Elem::U32 => "__vybe_ta_u32",
            Elem::F32 => "__vybe_ta_f32",
            Elem::F64 => "__vybe_ta_f64",
            Elem::BigI64 => "__vybe_ta_bi64",
            Elem::BigU64 => "__vybe_ta_bu64",
        }
    }

    /// Default zero-value for construction.
    fn zero(self) -> Value {
        match self {
            Elem::F32 | Elem::F64 => Value::F64(0.0),
            Elem::BigI64 | Elem::BigU64 => Value::I64(0),
            _ => Value::I32(0),
        }
    }

    /// Coerce an arbitrary input Value to this variant's element value,
    /// applying sign-extension / zero-extension / clamping per spec.
    fn coerce(self, v: &Value) -> Value {
        match self {
            Elem::I8 => Value::I32((v.as_i32() as i8) as i32),
            Elem::U8 => Value::I32((v.as_i32() & 0xFF) as i32),
            Elem::U8Clamped => {
                let n = v.as_f64();
                let clamped = if n.is_nan() { 0.0 } else { n.clamp(0.0, 255.0) };
                Value::I32(clamped.round() as i32)
            }
            Elem::I16 => Value::I32((v.as_i32() as i16) as i32),
            Elem::U16 => Value::I32((v.as_i32() & 0xFFFF) as i32),
            Elem::I32 => Value::I32(v.as_i32()),
            Elem::U32 => Value::I32(v.as_i32()),
            Elem::F32 => Value::F64((v.as_f64()) as f32 as f64),
            Elem::F64 => Value::F64(v.as_f64()),
            Elem::BigI64 => Value::I64(match v {
                Value::I64(n) => *n,
                _ => v.as_i32() as i64,
            }),
            Elem::BigU64 => Value::I64(match v {
                Value::I64(n) => *n,
                _ => v.as_i32() as i64,
            }),
        }
    }
}

/// Construct a new typed-array object of this variant with `length`
/// zero-initialised elements.
fn new_typed_array(variant: Elem, length: i32) -> Value {
    let zero = variant.zero();
    let elements: Vec<Value> = (0..length.max(0)).map(|_| zero.clone()).collect();
    let mut obj = Object::new_array(elements);
    obj.properties.insert(variant.tag().into(), Value::I32(1));
    obj.properties.insert("length".into(), Value::I32(length.max(0)));
    obj.properties.insert("byteLength".into(),
        Value::I32(length.max(0) * variant.bytes_per_element()));
    obj.properties.insert("byteOffset".into(), Value::I32(0));
    obj.properties.insert("BYTES_PER_ELEMENT".into(),
        Value::I32(variant.bytes_per_element()));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_typed(args: &[Value], idx: usize, variant: Elem) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(variant.tag()).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

pub fn register(vm: &mut VM) {
    use Elem::*;
    for variant in &[I8, U8, U8Clamped, I16, U16, I32, U32, F32, F64, BigI64, BigU64] {
        register_variant(vm, *variant);
    }
}

fn register_variant(vm: &mut VM, variant: Elem) {
    let module = variant.module();

    // ── Construction ─────────────────────────────────────────────────

    vm.register_host_fn(module, "newWithLength",
        Box::new(move |_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            new_typed_array(variant, n)
        }));

    vm.register_host_fn(module, "newFromBuffer",
        Box::new(move |_ctx, args| {
            // (buffer, byteOffset, length) — MVP: copy from buffer bytes,
            // re-interpreting as the variant's element type.
            let byte_offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0).max(0) as usize;
            let requested_len = args.get(2).map(|v| v.as_i32()).unwrap_or(-1);
            if let Some(Value::Object(buf)) = args.first() {
                let b = buf.lock().unwrap();
                if let ObjectKind::Array(ref bytes) = b.kind {
                    let bpe = variant.bytes_per_element() as usize;
                    let avail = bytes.len().saturating_sub(byte_offset);
                    let count = if requested_len < 0 { avail / bpe } else { requested_len as usize };
                    // For MVP: zero-fill; Phase B4 will re-interpret
                    // buffer bytes as typed-array elements with correct
                    // endianness (little-endian per spec).
                    return new_typed_array(variant, count as i32);
                }
            }
            new_typed_array(variant, 0)
        }));

    vm.register_host_fn(module, "newFromIterable",
        Box::new(move |_ctx, args| {
            // Source is an iterable; MVP handles Array source.
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref elems) = s.kind {
                    let coerced: Vec<Value> = elems.iter().map(|v| variant.coerce(v)).collect();
                    let len = coerced.len() as i32;
                    drop(s);
                    let mut obj = Object::new_array(coerced);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(len * variant.bytes_per_element()));
                    obj.properties.insert("byteOffset".into(), Value::I32(0));
                    obj.properties.insert("BYTES_PER_ELEMENT".into(),
                        Value::I32(variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    vm.register_host_fn(module, "newFromTypedArray",
        Box::new(move |_ctx, args| {
            // Copy elements from another typed array.
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref elems) = s.kind {
                    let coerced: Vec<Value> = elems.iter().map(|v| variant.coerce(v)).collect();
                    let len = coerced.len() as i32;
                    drop(s);
                    let mut obj = Object::new_array(coerced);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(len * variant.bytes_per_element()));
                    obj.properties.insert("byteOffset".into(), Value::I32(0));
                    obj.properties.insert("BYTES_PER_ELEMENT".into(),
                        Value::I32(variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    // Static from / of — aliases for the iterable-based constructor.
    vm.register_host_fn(module, "from",
        Box::new(move |_ctx, args| {
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref elems) = s.kind {
                    let coerced: Vec<Value> = elems.iter().map(|v| variant.coerce(v)).collect();
                    let len = coerced.len() as i32;
                    drop(s);
                    let mut obj = Object::new_array(coerced);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(len * variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    vm.register_host_fn(module, "of",
        Box::new(move |_ctx, args| {
            // `of(v1, v2, ..., vN)` — args are the elements directly.
            let coerced: Vec<Value> = args.iter().map(|v| variant.coerce(v)).collect();
            let len = coerced.len() as i32;
            let mut obj = Object::new_array(coerced);
            obj.properties.insert(variant.tag().into(), Value::I32(1));
            obj.properties.insert("length".into(), Value::I32(len));
            obj.properties.insert("byteLength".into(),
                Value::I32(len * variant.bytes_per_element()));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // ── Properties ───────────────────────────────────────────────────

    vm.register_host_fn(module, "buffer",
        Box::new(move |_ctx, args| {
            // We don't maintain a distinct ArrayBuffer for MVP — return
            // null to signal "no external buffer". Phase B4 will back
            // typed arrays with a real ArrayBuffer.
            if is_typed(args, 0, variant).is_some() {
                return Value::Null;
            }
            Value::Null
        }));

    vm.register_host_fn(module, "byteOffset",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                return o.properties.get("byteOffset").cloned().unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }));

    vm.register_host_fn(module, "byteLength",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    return Value::I32(elems.len() as i32 * variant.bytes_per_element());
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn(module, "length",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    return Value::I32(elems.len() as i32);
                }
            }
            Value::I32(0)
        }));

    // ── Element access ───────────────────────────────────────────────

    vm.register_host_fn(module, "get",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return variant.zero(); }
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    return elems.get(i as usize).cloned().unwrap_or_else(|| variant.zero());
                }
            }
            variant.zero()
        }));

    vm.register_host_fn(module, "at",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return variant.zero(); }
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    return elems.get(i as usize).cloned().unwrap_or_else(|| variant.zero());
                }
            }
            variant.zero()
        }));

    vm.register_host_fn(module, "set",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 { return Value::Null; }
            let val = args.get(2).cloned().unwrap_or(variant.zero());
            let coerced = variant.coerce(&val);
            if let Some(ta) = is_typed(args, 0, variant) {
                let mut o = ta.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    if let Some(slot) = elems.get_mut(i as usize) {
                        *slot = coerced;
                    }
                }
            }
            Value::Null
        }));

    vm.register_host_fn(module, "setArray",
        Box::new(move |_ctx, args| {
            // (ta, source, offset) — copy coerced elements from source
            let offset = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let source: Vec<Value> = match args.get(1) {
                Some(Value::Object(s)) => {
                    let sl = s.lock().unwrap();
                    if let ObjectKind::Array(ref e) = sl.kind {
                        e.iter().map(|v| variant.coerce(v)).collect()
                    } else { Vec::new() }
                }
                _ => Vec::new(),
            };
            if let Some(ta) = is_typed(args, 0, variant) {
                let mut o = ta.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    for (i, v) in source.into_iter().enumerate() {
                        if let Some(slot) = elems.get_mut(offset + i) {
                            *slot = v;
                        }
                    }
                }
            }
            Value::Null
        }));

    // ── Mutators that don't change length ────────────────────────────

    vm.register_host_fn(module, "copyWithin",
        Box::new(move |_ctx, args| {
            let target = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta) = is_typed(args, 0, variant) {
                let mut o = ta.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    let len = elems.len() as i32;
                    let t = target.max(0).min(len) as usize;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    let slice: Vec<Value> = elems[s..e].to_vec();
                    let max_copy = (len as usize - t).min(slice.len());
                    elems[t..t + max_copy].clone_from_slice(&slice[..max_copy]);
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "fill",
        Box::new(move |_ctx, args| {
            let val = args.get(1).map(|v| variant.coerce(v)).unwrap_or(variant.zero());
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta) = is_typed(args, 0, variant) {
                let mut o = ta.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    let len = elems.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    for i in s..e {
                        elems[i] = val.clone();
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "reverse",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let mut o = ta.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    elems.reverse();
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "toReversed",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let mut rev = elems.clone();
                    rev.reverse();
                    let len = rev.len() as i32;
                    let mut obj = Object::new_array(rev);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(len * variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    vm.register_host_fn(module, "sort",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let mut o = ta.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = o.kind {
                    elems.sort_by(|a, b| {
                        a.as_f64().partial_cmp(&b.as_f64())
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));

    vm.register_host_fn(module, "toSorted",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let mut sorted = elems.clone();
                    sorted.sort_by(|a, b| {
                        a.as_f64().partial_cmp(&b.as_f64())
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                    let len = sorted.len() as i32;
                    let mut obj = Object::new_array(sorted);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(len * variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    // ── Slicing ──────────────────────────────────────────────────────

    vm.register_host_fn(module, "slice",
        Box::new(move |_ctx, args| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let len = elems.len() as i32;
                    let s = (if start < 0 { len + start } else { start }).max(0).min(len) as usize;
                    let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                    let sub: Vec<Value> = if s < e { elems[s..e].to_vec() } else { Vec::new() };
                    let sub_len = sub.len() as i32;
                    let mut obj = Object::new_array(sub);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(sub_len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(sub_len * variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    vm.register_host_fn(module, "subarray",
        Box::new(move |_ctx, args| {
            // MVP: same semantics as slice; Phase B4 makes subarray
            // share storage with the parent via a proper view impl.
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let len = elems.len() as i32;
                    let s = (if start < 0 { len + start } else { start }).max(0).min(len) as usize;
                    let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                    let sub: Vec<Value> = if s < e { elems[s..e].to_vec() } else { Vec::new() };
                    let sub_len = sub.len() as i32;
                    let mut obj = Object::new_array(sub);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(sub_len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(sub_len * variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    // ── Search ───────────────────────────────────────────────────────

    vm.register_host_fn(module, "indexOf",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).map(|v| variant.coerce(v)).unwrap_or(variant.zero());
            let from = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let start = from.max(0) as usize;
                    for (i, v) in elems.iter().enumerate().skip(start) {
                        if v.eq(&needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }));

    vm.register_host_fn(module, "lastIndexOf",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).map(|v| variant.coerce(v)).unwrap_or(variant.zero());
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    for (i, v) in elems.iter().enumerate().rev() {
                        if v.eq(&needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }));

    vm.register_host_fn(module, "includes",
        Box::new(move |_ctx, args| {
            let needle = args.get(1).map(|v| variant.coerce(v)).unwrap_or(variant.zero());
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    for v in elems {
                        if v.eq(&needle) { return Value::I32(1); }
                    }
                }
            }
            Value::I32(0)
        }));

    // ── join / toString ──────────────────────────────────────────────

    vm.register_host_fn(module, "join",
        Box::new(move |_ctx, args| {
            let sep = args.get(1).map(|v| format!("{}", v)).unwrap_or_else(|| ",".into());
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let parts: Vec<String> = elems.iter().map(|e| format!("{}", e)).collect();
                    return Value::String(Arc::from(parts.join(&sep).as_str()));
                }
            }
            Value::String(Arc::from(""))
        }));

    vm.register_host_fn(module, "toString",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let parts: Vec<String> = elems.iter().map(|e| format!("{}", e)).collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }));

    vm.register_host_fn(module, "toLocaleString",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let parts: Vec<String> = elems.iter().map(|e| format!("{}", e)).collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }));

    // ── Iteration / keys/values/entries ──────────────────────────────

    vm.register_host_fn(module, "keys",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let out: Vec<Value> = (0..elems.len()).map(|i| Value::I32(i as i32)).collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(out))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn(module, "values",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn(module, "entries",
        Box::new(move |_ctx, args| {
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let pairs: Vec<Value> = elems.iter().enumerate()
                        .map(|(i, e)| Value::Object(Arc::new(Mutex::new(
                            Object::new_array(vec![Value::I32(i as i32), e.clone()])
                        ))))
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    // ── with(i, v) ───────────────────────────────────────────────────

    vm.register_host_fn(module, "with",
        Box::new(move |_ctx, args| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or(variant.zero());
            if let Some(ta) = is_typed(args, 0, variant) {
                let o = ta.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let len = elems.len() as i32;
                    let idx = if i < 0 { len + i } else { i };
                    if idx < 0 || idx >= len {
                        return args.first().cloned().unwrap_or(Value::Null);
                    }
                    let mut out = elems.clone();
                    out[idx as usize] = variant.coerce(&val);
                    let len = out.len() as i32;
                    let mut obj = Object::new_array(out);
                    obj.properties.insert(variant.tag().into(), Value::I32(1));
                    obj.properties.insert("length".into(), Value::I32(len));
                    obj.properties.insert("byteLength".into(),
                        Value::I32(len * variant.bytes_per_element()));
                    return Value::Object(Arc::new(Mutex::new(obj)));
                }
            }
            new_typed_array(variant, 0)
        }));

    // ── Higher-order callback methods ────────────────────────────────
    // (stubs — Phase B12 hooks up real callback dispatch)

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
                    "find" | "findLast" => variant.zero(),
                    "reduce" | "reduceRight" => args.get(2).cloned().unwrap_or(variant.zero()),
                    _ => args.first().cloned().unwrap_or(Value::Null),
                }
            }));
    }
}
