//! # `ecma:arraybuffer`, `ecma:sharedarraybuffer`, `ecma:dataview`
//!
//! Native Rust impls satisfying the imports declared in
//! `crates/vybe_runtime/src/wasm/js_arraybuffer_builtins.rs` per
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
use vybe_runtime::VM;
use vybe_runtime::value::{ArrayBufferState, Object, ObjectKind, Value};

pub const DV_TAG: &str = "__vybe_js_dataview";
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
    obj.properties
        .insert("byteLength".into(), Value::I32(n as i32));
    obj.properties
        .insert("maxByteLength".into(), Value::I32(max as i32));
    obj.properties
        .insert("resizable".into(), Value::Bool(resizable));
    obj.properties.insert("detached".into(), Value::Bool(false));
    let type_name = if shared {
        "SharedArrayBuffer"
    } else {
        "ArrayBuffer"
    };
    obj.properties
        .insert("__type".into(), Value::String(Arc::from(type_name)));
    Value::Object(vybe_runtime::heap::alloc(obj))
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

fn apply_arraybuffer_receiver_species(result: &Value, receiver: &Arc<Mutex<Object>>) {
    let Some(ctor) = crate::object::proto_walk_get(receiver, "constructor") else {
        return;
    };
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
    if !matches!(result_lock.kind, ObjectKind::ArrayBuffer(_)) {
        return;
    }
    if let Some(proto) = prototype {
        result_lock
            .properties
            .insert("__proto__".into(), Value::Object(proto));
    }
    if let Some(name) = name {
        let types = vybe_runtime::heap::alloc(Object::new_array(vec![Value::String(
            Arc::from(name.as_str()),
        )]));
        result_lock
            .properties
            .insert("__types".into(), Value::Object(types));
    }
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
    vm.register_host_fn(
        "ecma:arraybuffer",
        "new",
        Box::new(|ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if n < 0 {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid ArrayBuffer length",
                ));
                return Value::Undefined;
            }
            let max = args
                .get(1)
                .and_then(|options| {
                    let Value::Object(obj) = options else {
                        return None;
                    };
                    obj.lock()
                        .unwrap()
                        .properties
                        .get("maxByteLength")
                        .map(|value| value.as_i32())
                })
                .unwrap_or(n);
            if max < n {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "ArrayBuffer maxByteLength is smaller than byteLength",
                ));
                return Value::Undefined;
            }
            new_arraybuffer(n, max, max != n, false)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "newWithLength",
        Box::new(|ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if n < 0 {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid ArrayBuffer length",
                ));
                return Value::Undefined;
            }
            new_arraybuffer(n, n, false, false)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "newResizable",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_i32()).unwrap_or(n);
            new_arraybuffer(n, max, true, false)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "byteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                return Value::I32(ab_byte_length_of(&ab) as i32);
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "maxByteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(state.max_byte_length as i32);
                }
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "resizable",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::Bool(state.resizable);
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "detached",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::Bool(state.detached);
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "slice",
        Box::new(|ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    if state.detached {
                        drop(o);
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "TypeError",
                            "ArrayBuffer is detached",
                        ));
                        return Value::Undefined;
                    }
                    let src = state.bytes.lock().unwrap();
                    let len = src.len() as i32;
                    let s = if start < 0 {
                        (len + start).max(0)
                    } else {
                        start.min(len)
                    } as usize;
                    let e = if end == i32::MAX {
                        len as usize
                    } else if end < 0 {
                        (len + end).max(0) as usize
                    } else {
                        end.min(len) as usize
                    };
                    let slice: Vec<u8> = if s < e {
                        src[s..e].to_vec()
                    } else {
                        Vec::new()
                    };
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
                    new_obj
                        .properties
                        .insert("byteLength".into(), Value::I32(slice_len as i32));
                    new_obj
                        .properties
                        .insert("maxByteLength".into(), Value::I32(slice_len as i32));
                    let out = Value::Object(vybe_runtime::heap::alloc(new_obj));
                    apply_arraybuffer_receiver_species(&out, &ab);
                    return out;
                }
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "resize",
        Box::new(|ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let new_len = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
                let mut o = ab.lock().unwrap();
                let ObjectKind::ArrayBuffer(ref mut state) = o.kind else {
                    return Value::Null;
                };
                if state.detached {
                    drop(o);
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "ArrayBuffer is detached",
                    ));
                    return Value::Undefined;
                }
                if !state.resizable {
                    drop(o);
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "ArrayBuffer is not resizable",
                    ));
                    return Value::Undefined;
                }
                if new_len > state.max_byte_length {
                    drop(o);
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        "ArrayBuffer resize exceeds maxByteLength",
                    ));
                    return Value::Undefined;
                }
                let mut bytes = state.bytes.lock().unwrap();
                bytes.resize(new_len, 0);
                drop(bytes);
                o.properties
                    .insert("byteLength".into(), Value::I32(new_len as i32));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "transfer",
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
                o.properties.insert("detached".into(), Value::Bool(true));
                drop(o);

                let target_len = if requested < 0 {
                    taken_bytes.len()
                } else {
                    requested.max(0) as usize
                };
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
                new_obj
                    .properties
                    .insert("byteLength".into(), Value::I32(target_len as i32));
                new_obj
                    .properties
                    .insert("maxByteLength".into(), Value::I32(target_len as i32));
                new_obj
                    .properties
                    .insert("resizable".into(), Value::Bool(false));
                new_obj
                    .properties
                    .insert("detached".into(), Value::Bool(false));
                return Value::Object(vybe_runtime::heap::alloc(new_obj));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "transferToFixedLength",
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
                o.properties.insert("detached".into(), Value::Bool(true));
                drop(o);

                let target_len = if requested < 0 {
                    taken_bytes.len()
                } else {
                    requested.max(0) as usize
                };
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
                new_obj
                    .properties
                    .insert("byteLength".into(), Value::I32(target_len as i32));
                new_obj
                    .properties
                    .insert("maxByteLength".into(), Value::I32(target_len as i32));
                new_obj
                    .properties
                    .insert("resizable".into(), Value::Bool(false));
                new_obj
                    .properties
                    .insert("detached".into(), Value::Bool(false));
                return Value::Object(vybe_runtime::heap::alloc(new_obj));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:arraybuffer",
        "isView",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if o.properties.get(DV_TAG).is_some() {
                    return Value::Bool(true);
                }
                if matches!(o.kind, ObjectKind::TypedArray(_)) {
                    return Value::Bool(true);
                }
            }
            Value::Bool(false)
        }),
    );
}

// ── SharedArrayBuffer ─────────────────────────────────────────────────

fn register_sharedarraybuffer(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "new",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            new_arraybuffer(n, n, false, true)
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "newWithLength",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            new_arraybuffer(n, n, false, true)
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "newGrowable",
        Box::new(|_ctx, args| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let max = args.get(1).map(|v| v.as_i32()).unwrap_or(n);
            new_arraybuffer(n, max, true, true)
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "byteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                return Value::I32(ab_byte_length_of(&ab) as i32);
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "maxByteLength",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::I32(state.max_byte_length as i32);
                }
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "growable",
        Box::new(|_ctx, args| {
            if let Some(ab) = is_arraybuffer(args, 0) {
                let o = ab.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                    return Value::Bool(state.resizable);
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "slice",
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
                    let slice: Vec<u8> = if s < e {
                        src[s..e].to_vec()
                    } else {
                        Vec::new()
                    };
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
                    new_obj
                        .properties
                        .insert("byteLength".into(), Value::I32(slice_len as i32));
                    new_obj
                        .properties
                        .insert("maxByteLength".into(), Value::I32(slice_len as i32));
                    return Value::Object(vybe_runtime::heap::alloc(new_obj));
                }
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:sharedarraybuffer",
        "grow",
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
                    o.properties
                        .insert("byteLength".into(), Value::I32(new_byte_len));
                }
            }
            Value::Null
        }),
    );
}

// ── DataView ──────────────────────────────────────────────────────────

pub fn new_dataview(buffer: Value, byte_offset: i32, byte_length: i32) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(DV_TAG.into(), Value::I32(1));
    obj.properties.insert(DV_BUFFER_PROP.into(), buffer.clone());
    obj.properties.insert("buffer".into(), buffer);
    obj.properties
        .insert(DV_OFFSET_PROP.into(), Value::I32(byte_offset.max(0)));
    obj.properties
        .insert(DV_LENGTH_PROP.into(), Value::I32(byte_length.max(0)));
    obj.properties
        .insert("byteOffset".into(), Value::I32(byte_offset.max(0)));
    obj.properties
        .insert("byteLength".into(), Value::I32(byte_length.max(0)));
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("DataView")));
    Value::Object(vybe_runtime::heap::alloc(obj))
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

/// §25.3.1.1/.2 bounds predicate: is a `count`-byte access at `offset` fully
/// inside the view? A negative offset is out of bounds (matches ToIndex
/// §7.1.22 rejecting negatives). Used to raise RangeError before a get/set.
fn dv_in_bounds(dv: &Arc<Mutex<Object>>, offset: i32, count: usize) -> bool {
    let Some((bytes_arc, base, view_len)) = dv_resolve(dv) else {
        return false;
    };
    if offset < 0 || (offset as usize + count) > view_len {
        return false;
    }
    let bytes = bytes_arc.lock().unwrap();
    base + offset as usize + count <= bytes.len()
}

fn dv_write_bytes(dv: &Arc<Mutex<Object>>, offset: i32, payload: &[u8]) -> bool {
    let Some((bytes_arc, base, view_len)) = dv_resolve(dv) else {
        return false;
    };
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
    vm.register_host_fn(
        "ecma:dataview",
        "new",
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
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "buffer",
        Box::new(|_ctx, args| {
            if let Some(dv) = is_dataview(args, 0) {
                let o = dv.lock().unwrap();
                return o
                    .properties
                    .get(DV_BUFFER_PROP)
                    .cloned()
                    .unwrap_or(Value::Null);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "byteOffset",
        Box::new(|_ctx, args| {
            if let Some(dv) = is_dataview(args, 0) {
                let o = dv.lock().unwrap();
                return o
                    .properties
                    .get(DV_OFFSET_PROP)
                    .cloned()
                    .unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "byteLength",
        Box::new(|_ctx, args| {
            if let Some(dv) = is_dataview(args, 0) {
                let o = dv.lock().unwrap();
                return o
                    .properties
                    .get(DV_LENGTH_PROP)
                    .cloned()
                    .unwrap_or(Value::I32(0));
            }
            Value::I32(0)
        }),
    );

    // Single-byte getters/setters (no endianness operand per spec)

    vm.register_host_fn(
        "ecma:dataview",
        "getInt8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                if let Some(bytes) = dv_read_bytes(&dv, offset, 1) {
                    return Value::I32(bytes[0] as i8 as i32);
                }
            }
            Value::I32(0)
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "getUint8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                if let Some(bytes) = dv_read_bytes(&dv, offset, 1) {
                    return Value::I32(bytes[0] as i32);
                }
            }
            Value::I32(0)
        }),
    );

    macro_rules! getter_multibyte {
        ($name:literal, $count:expr, $ty:ty, $le:path, $be:path, $wrap:expr) => {
            vm.register_host_fn(
                "ecma:dataview",
                $name,
                Box::new(|ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    let little_endian = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
                    if let Some(dv) = is_dataview(args, 0) {
                        if let Some(bytes) = dv_read_bytes(&dv, offset, $count) {
                            let mut arr = [0u8; $count];
                            arr.copy_from_slice(&bytes);
                            let val: $ty = if little_endian { $le(arr) } else { $be(arr) };
                            return $wrap(val);
                        }
                        // §25.3.1.1 GetViewValue step 8: getIndex + elementSize
                        // > viewSize → RangeError. A negative offset is also a
                        // RangeError via ToIndex (§7.1.22). `dv_read_bytes`
                        // returns None for both.
                        let err = crate::error::new_error(
                            ctx,
                            "RangeError",
                            "Offset is outside the bounds of the DataView",
                        );
                        ctx.throw_value(err);
                        return Value::Undefined;
                    }
                    $wrap(<$ty>::default())
                }),
            );
        };
    }

    getter_multibyte!(
        "getInt16",
        2,
        i16,
        i16::from_le_bytes,
        i16::from_be_bytes,
        |v| Value::I32(v as i32)
    );
    getter_multibyte!(
        "getUint16",
        2,
        u16,
        u16::from_le_bytes,
        u16::from_be_bytes,
        |v| Value::I32(v as i32)
    );
    getter_multibyte!(
        "getInt32",
        4,
        i32,
        i32::from_le_bytes,
        i32::from_be_bytes,
        |v| Value::I32(v)
    );
    getter_multibyte!(
        "getUint32",
        4,
        u32,
        u32::from_le_bytes,
        u32::from_be_bytes,
        |v| Value::I32(v as i32)
    );
    // §25.3.4: the BigInt64/BigUint64 accessors traffic in BigInt values.
    getter_multibyte!(
        "getBigInt64",
        8,
        i64,
        i64::from_le_bytes,
        i64::from_be_bytes,
        |v| Value::bigint_i64(v)
    );
    getter_multibyte!(
        "getBigUint64",
        8,
        u64,
        u64::from_le_bytes,
        u64::from_be_bytes,
        |v| Value::bigint_u64(v)
    );
    getter_multibyte!(
        "getFloat32",
        4,
        f32,
        f32::from_le_bytes,
        f32::from_be_bytes,
        |v| Value::F64(v as f64)
    );
    getter_multibyte!(
        "getFloat64",
        8,
        f64,
        f64::from_le_bytes,
        f64::from_be_bytes,
        |v| Value::F64(v)
    );

    vm.register_host_fn(
        "ecma:dataview",
        "setInt8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                dv_write_bytes(&dv, offset, &[(val as i8) as u8]);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "setUint8",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(dv) = is_dataview(args, 0) {
                dv_write_bytes(&dv, offset, &[val as u8]);
            }
            Value::Null
        }),
    );

    macro_rules! setter_multibyte {
        ($name:literal, $count:expr, $val_extract:expr, $ty:ty, $le:ident, $be:ident) => {
            vm.register_host_fn(
                "ecma:dataview",
                $name,
                Box::new(|ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    let val: $ty = $val_extract(args.get(2));
                    let little_endian = args.get(3).map(|v| v.as_i32()).unwrap_or(0) != 0;
                    let bytes = if little_endian { val.$le() } else { val.$be() };
                    if let Some(dv) = is_dataview(args, 0) {
                        // §25.3.1.2 SetViewValue: out-of-bounds (or negative
                        // offset via ToIndex) → RangeError.
                        if !dv_write_bytes(&dv, offset, &bytes) {
                            let err = crate::error::new_error(
                                ctx,
                                "RangeError",
                                "Offset is outside the bounds of the DataView",
                            );
                            ctx.throw_value(err);
                        }
                    }
                    Value::Null
                }),
            );
        };
    }

    setter_multibyte!(
        "setInt16",
        2,
        |v: Option<&Value>| v.map(|x| x.as_i32() as i16).unwrap_or(0),
        i16,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setUint16",
        2,
        |v: Option<&Value>| v.map(|x| x.as_i32() as u16).unwrap_or(0),
        u16,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setInt32",
        4,
        |v: Option<&Value>| v.map(|x| x.as_i32()).unwrap_or(0),
        i32,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setUint32",
        4,
        |v: Option<&Value>| v.map(|x| x.as_i32() as u32).unwrap_or(0),
        u32,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setBigInt64",
        8,
        |v: Option<&Value>| v
            .map(|x| match x {
                Value::BigInt(n) => n.to_i64_wrapping(),
                Value::I64(n) => *n,
                other => other.as_i32() as i64,
            })
            .unwrap_or(0),
        i64,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setBigUint64",
        8,
        |v: Option<&Value>| v
            .map(|x| match x {
                Value::BigInt(n) => n.to_u64_wrapping(),
                Value::I64(n) => *n as u64,
                other => other.as_i32() as u64,
            })
            .unwrap_or(0),
        u64,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setFloat32",
        4,
        |v: Option<&Value>| v.map(|x| x.as_f64() as f32).unwrap_or(0.0),
        f32,
        to_le_bytes,
        to_be_bytes
    );
    setter_multibyte!(
        "setFloat64",
        8,
        |v: Option<&Value>| v.map(|x| x.as_f64()).unwrap_or(0.0),
        f64,
        to_le_bytes,
        to_be_bytes
    );

    // ── Named constructors ────────────────────────────────────────────────────

    vm.register_host_fn(
        "ecma:dataview",
        "newWithOffset",
        Box::new(|_ctx, args| {
            let buffer = args.first().cloned().unwrap_or(Value::Null);
            let byte_offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
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
            let byte_length = (buffer_len - byte_offset).max(0);
            new_dataview(buffer, byte_offset, byte_length)
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "newWithOffsetAndLength",
        Box::new(|_ctx, args| {
            let buffer = args.first().cloned().unwrap_or(Value::Null);
            let byte_offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let byte_length = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            new_dataview(buffer, byte_offset, byte_length)
        }),
    );

    // ── LE/BE aliased getters and setters ────────────────────────────────────

    macro_rules! getter_le {
        ($name:literal, $count:expr, $ty:ty, $wrap:expr) => {
            vm.register_host_fn(
                "ecma:dataview",
                $name,
                Box::new(|_ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    if let Some(dv) = is_dataview(args, 0) {
                        if let Some(bytes) = dv_read_bytes(&dv, offset, $count) {
                            let mut arr = [0u8; $count];
                            arr.copy_from_slice(&bytes);
                            let val: $ty = <$ty>::from_le_bytes(arr);
                            return $wrap(val);
                        }
                    }
                    $wrap(<$ty>::default())
                }),
            );
        };
    }

    macro_rules! getter_be {
        ($name:literal, $count:expr, $ty:ty, $wrap:expr) => {
            vm.register_host_fn(
                "ecma:dataview",
                $name,
                Box::new(|_ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    if let Some(dv) = is_dataview(args, 0) {
                        if let Some(bytes) = dv_read_bytes(&dv, offset, $count) {
                            let mut arr = [0u8; $count];
                            arr.copy_from_slice(&bytes);
                            let val: $ty = <$ty>::from_be_bytes(arr);
                            return $wrap(val);
                        }
                    }
                    $wrap(<$ty>::default())
                }),
            );
        };
    }

    macro_rules! setter_le {
        ($name:literal, $count:expr, $val_extract:expr, $ty:ty) => {
            vm.register_host_fn(
                "ecma:dataview",
                $name,
                Box::new(|_ctx, args| {
                    let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    let val: $ty = $val_extract(args.get(2));
                    if let Some(dv) = is_dataview(args, 0) {
                        dv_write_bytes(&dv, offset, &val.to_le_bytes());
                    }
                    Value::Null
                }),
            );
        };
    }

    getter_le!("getInt16LE", 2, i16, |v| Value::I32(v as i32));
    getter_be!("getInt16BE", 2, i16, |v| Value::I32(v as i32));
    getter_le!("getInt32LE", 4, i32, |v| Value::I32(v));
    getter_le!("getUint32LE", 4, u32, |v| Value::F64(v as f64));
    getter_le!("getFloat32LE", 4, f32, |v| Value::F64(v as f64));
    getter_le!("getFloat64LE", 8, f64, |v| Value::F64(v));
    getter_le!("getBigInt64LE", 8, i64, |v| Value::I64(v));
    getter_le!("getBigUint64LE", 8, u64, |v| Value::I64(v as i64));

    setter_le!(
        "setInt16LE",
        2,
        |v: Option<&Value>| v.map(|x| x.as_i32() as i16).unwrap_or(0),
        i16
    );
    setter_le!(
        "setInt32LE",
        4,
        |v: Option<&Value>| v.map(|x| x.as_i32()).unwrap_or(0),
        i32
    );
    setter_le!(
        "setUint32LE",
        4,
        |v: Option<&Value>| v.map(|x| x.as_f64() as u32).unwrap_or(0),
        u32
    );
    setter_le!(
        "setFloat32LE",
        4,
        |v: Option<&Value>| v.map(|x| x.as_f64() as f32).unwrap_or(0.0),
        f32
    );
    setter_le!(
        "setFloat64LE",
        8,
        |v: Option<&Value>| v.map(|x| x.as_f64()).unwrap_or(0.0),
        f64
    );
    setter_le!(
        "setBigInt64LE",
        8,
        |v: Option<&Value>| v
            .map(|x| match x {
                Value::I64(n) => *n,
                other => other.as_i32() as i64,
            })
            .unwrap_or(0),
        i64
    );
    setter_le!(
        "setBigUint64LE",
        8,
        |v: Option<&Value>| v
            .map(|x| match x {
                Value::I64(n) => *n as u64,
                other => other.as_i32() as u64,
            })
            .unwrap_or(0),
        u64
    );

    // ── Float16 (ES2025 §25.3.4.*) ───────────────────────────────────────────

    vm.register_host_fn(
        "ecma:dataview",
        "setFloat16",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).map(|v| v.as_f64()).unwrap_or(0.0);
            let little_endian = args.get(3).map(|v| v.as_bool()).unwrap_or(false);
            let bits = f64_to_f16(val);
            let bytes = if little_endian {
                bits.to_le_bytes()
            } else {
                bits.to_be_bytes()
            };
            if let Some(dv) = is_dataview(args, 0) {
                dv_write_bytes(&dv, offset, &bytes);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "ecma:dataview",
        "getFloat16",
        Box::new(|_ctx, args| {
            let offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let little_endian = args.get(2).map(|v| v.as_bool()).unwrap_or(false);
            if let Some(dv) = is_dataview(args, 0) {
                if let Some(bytes) = dv_read_bytes(&dv, offset, 2) {
                    let mut arr = [0u8; 2];
                    arr.copy_from_slice(&bytes);
                    let bits = if little_endian {
                        u16::from_le_bytes(arr)
                    } else {
                        u16::from_be_bytes(arr)
                    };
                    return Value::F64(f16_to_f64(bits));
                }
            }
            Value::Undefined
        }),
    );
}

fn f64_to_f16(v: f64) -> u16 {
    if v.is_nan() {
        return 0x7E00;
    }
    let bits = v.to_bits();
    let sign = ((bits >> 63) as u16) << 15;
    let exp = ((bits >> 52) & 0x7FF) as i32 - 1023;
    let mant = (bits & 0x000F_FFFF_FFFF_FFFF) >> 42;
    if exp == 1024 {
        return sign
            | 0x7C00
            | (if (bits & 0x000F_FFFF_FFFF_FFFF) != 0 {
                0x0200
            } else {
                0
            });
    }
    let f16_exp = exp + 15;
    if f16_exp <= 0 {
        if f16_exp < -10 {
            return sign;
        }
        let shifted = (1u64 << 10 | mant) >> (1 - f16_exp);
        return sign | (shifted as u16);
    }
    if f16_exp >= 31 {
        return sign | 0x7C00;
    }
    sign | ((f16_exp as u16) << 10) | (mant as u16)
}

fn f16_to_f64(bits: u16) -> f64 {
    let sign = ((bits >> 15) & 1) as u64;
    let exp = ((bits >> 10) & 0x1F) as i32;
    let mant = (bits & 0x3FF) as u64;
    if exp == 31 {
        return f64::from_bits((sign << 63) | 0x7FF0_0000_0000_0000 | (mant << 42));
    }
    if exp == 0 {
        if mant == 0 {
            return if sign == 0 { 0.0 } else { -0.0 };
        }
        let v = (mant as f64) * (1.0 / (1024.0 * 16384.0));
        return if sign == 0 { v } else { -v };
    }
    let f64_exp = (exp - 15 + 1023) as u64;
    f64::from_bits((sign << 63) | (f64_exp << 52) | (mant << 42))
}

// ── Public method dispatch — called from ecma::value ─────────────────────

/// Dispatches instance method calls on ArrayBuffer objects.
/// `args[0]` = the ArrayBuffer object; remaining args are user-supplied.
pub fn dispatch_arraybuffer_method(
    ctx: &mut vybe_runtime::HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Option<Value> {
    match method {
        "slice" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::ArrayBuffer(ref state) = o.kind {
                if state.detached {
                    drop(o);
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "ArrayBuffer is detached",
                    ));
                    return Some(Value::Undefined);
                }
                let src = state.bytes.lock().unwrap();
                let len = src.len() as i32;
                let start = args.first().map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(1).map(|v| v.as_i32()).unwrap_or(len);
                let s = (if start < 0 { len + start } else { start })
                    .max(0)
                    .min(len) as usize;
                let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                let slice: Vec<u8> = if s < e {
                    src[s..e].to_vec()
                } else {
                    Vec::new()
                };
                drop(src);
                let slice_len = slice.len();
                let shared = state.shared;
                drop(o);
                let new_state = ArrayBufferState {
                    bytes: Arc::new(Mutex::new(slice)),
                    max_byte_length: slice_len,
                    resizable: false,
                    detached: false,
                    shared,
                };
                let type_name = if shared {
                    "SharedArrayBuffer"
                } else {
                    "ArrayBuffer"
                };
                let mut new_obj = Object::new();
                new_obj.kind = ObjectKind::ArrayBuffer(new_state);
                new_obj
                    .properties
                    .insert("byteLength".into(), Value::I32(slice_len as i32));
                new_obj
                    .properties
                    .insert("maxByteLength".into(), Value::I32(slice_len as i32));
                new_obj
                    .properties
                    .insert("__type".into(), Value::String(Arc::from(type_name)));
                let out = Value::Object(vybe_runtime::heap::alloc(new_obj));
                apply_arraybuffer_receiver_species(&out, &obj);
                return Some(out);
            }
            Some(Value::Null)
        }
        "resize" => {
            let new_len = args
                .first()
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            let mut o = obj.lock().unwrap();
            let ObjectKind::ArrayBuffer(ref mut state) = o.kind else {
                return Some(Value::Null);
            };
            if state.detached {
                drop(o);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "ArrayBuffer is detached",
                ));
                return Some(Value::Undefined);
            }
            if !state.resizable {
                drop(o);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "ArrayBuffer is not resizable",
                ));
                return Some(Value::Undefined);
            }
            if new_len > state.max_byte_length {
                drop(o);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "ArrayBuffer resize exceeds maxByteLength",
                ));
                return Some(Value::Undefined);
            }
            let mut bytes = state.bytes.lock().unwrap();
            bytes.resize(new_len, 0);
            drop(bytes);
            o.properties
                .insert("byteLength".into(), Value::I32(new_len as i32));
            Some(Value::Undefined)
        }
        "transfer" | "transferToFixedLength" => {
            let requested = args.first().map(|v| v.as_i32()).unwrap_or(-1);
            let mut o = obj.lock().unwrap();
            let (taken_bytes, shared) = if let ObjectKind::ArrayBuffer(ref mut state) = o.kind {
                if state.detached {
                    drop(o);
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "ArrayBuffer is detached",
                    ));
                    return Some(Value::Undefined);
                }
                let mut src = state.bytes.lock().unwrap();
                let taken = std::mem::take(&mut *src);
                drop(src);
                state.detached = true;
                (taken, state.shared)
            } else {
                return Some(Value::Null);
            };
            o.properties.insert("byteLength".into(), Value::I32(0));
            o.properties.insert("detached".into(), Value::Bool(true));
            drop(o);

            let target_len = if requested < 0 {
                taken_bytes.len()
            } else {
                requested.max(0) as usize
            };
            let mut new_bytes = taken_bytes;
            new_bytes.resize(target_len, 0);
            let new_state = ArrayBufferState {
                bytes: Arc::new(Mutex::new(new_bytes)),
                max_byte_length: target_len,
                resizable: false,
                detached: false,
                shared,
            };
            let type_name = if shared {
                "SharedArrayBuffer"
            } else {
                "ArrayBuffer"
            };
            let mut new_obj = Object::new();
            new_obj.kind = ObjectKind::ArrayBuffer(new_state);
            new_obj
                .properties
                .insert("byteLength".into(), Value::I32(target_len as i32));
            new_obj
                .properties
                .insert("maxByteLength".into(), Value::I32(target_len as i32));
            new_obj
                .properties
                .insert("resizable".into(), Value::Bool(false));
            new_obj
                .properties
                .insert("detached".into(), Value::Bool(false));
            new_obj
                .properties
                .insert("__type".into(), Value::String(Arc::from(type_name)));
            Some(Value::Object(vybe_runtime::heap::alloc(new_obj)))
        }
        _ => None,
    }
}

/// Dispatches instance method calls on DataView objects.
/// `args[0]` = the DataView object; remaining args are user-supplied.
pub fn dispatch_dataview_method(
    ctx: &mut vybe_runtime::HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Option<Value> {
    // §25.3.1.1 GetViewValue / §25.3.1.2 SetViewValue: bounds-check the
    // access up front — `getIndex + elementSize > viewSize`, or a negative
    // offset (via ToIndex §7.1.22), is a RangeError. Done once here so the
    // per-method arms below need no per-call guard.
    let elem_size = match method {
        "getInt8" | "getUint8" | "setInt8" | "setUint8" => Some(1),
        "getInt16" | "getUint16" | "setInt16" | "setUint16" => Some(2),
        "getInt32" | "getUint32" | "getFloat32" | "setInt32" | "setUint32" | "setFloat32" => {
            Some(4)
        }
        "getFloat64" | "getBigInt64" | "getBigUint64" | "setFloat64" | "setBigInt64"
        | "setBigUint64" => Some(8),
        _ => None,
    };
    if let Some(size) = elem_size {
        let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
        if !dv_in_bounds(&obj, offset, size) {
            let err = crate::error::new_error(
                ctx,
                "RangeError",
                "Offset is outside the bounds of the DataView",
            );
            ctx.throw_value(err);
            return Some(Value::Undefined);
        }
    }
    match method {
        "getInt8" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if let Some(bytes) = dv_read_bytes(&obj, offset, 1) {
                return Some(Value::I32(bytes[0] as i8 as i32));
            }
            Some(Value::I32(0))
        }
        "getUint8" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if let Some(bytes) = dv_read_bytes(&obj, offset, 1) {
                return Some(Value::I32(bytes[0] as i32));
            }
            Some(Value::I32(0))
        }
        "getInt16" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 2) {
                let arr = [bytes[0], bytes[1]];
                let v = if le {
                    i16::from_le_bytes(arr)
                } else {
                    i16::from_be_bytes(arr)
                };
                return Some(Value::I32(v as i32));
            }
            Some(Value::I32(0))
        }
        "getUint16" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 2) {
                let arr = [bytes[0], bytes[1]];
                let v = if le {
                    u16::from_le_bytes(arr)
                } else {
                    u16::from_be_bytes(arr)
                };
                return Some(Value::I32(v as i32));
            }
            Some(Value::I32(0))
        }
        "getInt32" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 4) {
                let arr = [bytes[0], bytes[1], bytes[2], bytes[3]];
                let v = if le {
                    i32::from_le_bytes(arr)
                } else {
                    i32::from_be_bytes(arr)
                };
                return Some(Value::I32(v));
            }
            Some(Value::I32(0))
        }
        "getUint32" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 4) {
                let arr = [bytes[0], bytes[1], bytes[2], bytes[3]];
                let v = if le {
                    u32::from_le_bytes(arr)
                } else {
                    u32::from_be_bytes(arr)
                };
                // Return as F64 to preserve full u32 range (JS numbers are f64)
                return Some(Value::F64(v as f64));
            }
            Some(Value::F64(0.0))
        }
        "getFloat32" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 4) {
                let arr = [bytes[0], bytes[1], bytes[2], bytes[3]];
                let v = if le {
                    f32::from_le_bytes(arr)
                } else {
                    f32::from_be_bytes(arr)
                };
                return Some(Value::F64(v as f64));
            }
            Some(Value::F64(0.0))
        }
        "getFloat64" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 8) {
                let arr = [
                    bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
                ];
                let v = if le {
                    f64::from_le_bytes(arr)
                } else {
                    f64::from_be_bytes(arr)
                };
                return Some(Value::F64(v));
            }
            Some(Value::F64(0.0))
        }
        "getBigInt64" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 8) {
                let arr = [
                    bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
                ];
                let v = if le {
                    i64::from_le_bytes(arr)
                } else {
                    i64::from_be_bytes(arr)
                };
                return Some(Value::bigint_i64(v));
            }
            Some(Value::bigint_i64(0))
        }
        "getBigUint64" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(1).map(|v| v.as_i32()).unwrap_or(0) != 0;
            if let Some(bytes) = dv_read_bytes(&obj, offset, 8) {
                let arr = [
                    bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
                ];
                let v = if le {
                    u64::from_le_bytes(arr)
                } else {
                    u64::from_be_bytes(arr)
                };
                return Some(Value::bigint_u64(v));
            }
            Some(Value::bigint_i64(0))
        }
        "setInt8" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            dv_write_bytes(&obj, offset, &[(val as i8) as u8]);
            Some(Value::Undefined)
        }
        "setUint8" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            dv_write_bytes(&obj, offset, &[val as u8]);
            Some(Value::Undefined)
        }
        "setInt16" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_i32() as i16).unwrap_or(0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        "setUint16" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_i32() as u16).unwrap_or(0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        "setInt32" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        "setUint32" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        "setFloat32" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_f64() as f32).unwrap_or(0.0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        "setFloat64" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        "setBigInt64" | "setBigUint64" => {
            let offset = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let val = args
                .get(1)
                .map(|x| match x {
                    Value::BigInt(n) => n.to_i64_wrapping(),
                    Value::I64(n) => *n,
                    other => other.as_i32() as i64,
                })
                .unwrap_or(0);
            let le = args.get(2).map(|v| v.as_i32()).unwrap_or(0) != 0;
            let bytes = if le {
                val.to_le_bytes()
            } else {
                val.to_be_bytes()
            };
            dv_write_bytes(&obj, offset, &bytes);
            Some(Value::Undefined)
        }
        _ => None,
    }
}
