//! # `wasm:js-structured-clone` host handler
//!
//! Implements the HTML structured-clone algorithm: deep-copy across
//! Array / Object / Map / Set / ArrayBuffer / TypedArray / DataView /
//! WeakMap / WeakSet (weak containers aren't cloneable per spec —
//! return as-is / share reference) / primitives.
//!
//! Cycle handling: uses a visited map keyed by `Arc` identity so
//! cyclic object graphs round-trip (unlike JSON).
//!
//! Spec: <https://html.spec.whatwg.org/multipage/structured-data.html#structured-clone>

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{
    ArrayBufferState, Object, ObjectKind, TypedArrayState, Value,
};
use vybe_bytecode::{HostContext, VM};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("wasm:js-structured-clone", "clone",
        Box::new(|_ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let mut seen: HashMap<usize, Value> = HashMap::new();
            deep_clone(&value, &mut seen)
        }));
}

fn deep_clone(v: &Value, seen: &mut HashMap<usize, Value>) -> Value {
    match v {
        // Primitives — pass through; they're Copy or immutable.
        Value::Null | Value::Undefined => v.clone(),
        Value::Bool(_) | Value::I32(_) | Value::I64(_) | Value::F64(_) => v.clone(),
        Value::BigInt(_) | Value::V128(_) => v.clone(),

        // Strings: Arc<str> is already cheap to clone; structured
        // clone of a string = the same string per spec.
        Value::String(_) => v.clone(),

        // Symbols: not cloneable per spec (structured clone throws
        // DataCloneError). MVP passes through — the compiler layer
        // can enforce the throw when we have exception dispatch.
        Value::Symbol(_) => v.clone(),

        // WeakRef: structured clone of a weak reference yields a
        // dead weak reference per spec. Pass through.
        Value::WeakRef(_) => v.clone(),

        Value::Object(obj) => clone_object(obj, seen),
    }
}

fn clone_object(obj: &Arc<Mutex<Object>>, seen: &mut HashMap<usize, Value>) -> Value {
    let id = Arc::as_ptr(obj) as usize;
    if let Some(already) = seen.get(&id) {
        return already.clone();
    }

    // Determine the kind first so we can place the freshly-allocated
    // target into `seen` before recursing — needed for cycle handling.
    let kind_tag = {
        let o = obj.lock().unwrap();
        kind_discriminant(&o.kind)
    };

    match kind_tag {
        KindTag::Ordinary => clone_ordinary(obj, id, seen),
        KindTag::Array => clone_array(obj, id, seen),
        KindTag::Map => clone_map(obj, id, seen),
        KindTag::Set => clone_set(obj, id, seen),
        KindTag::ArrayBuffer => clone_arraybuffer(obj, id, seen),
        KindTag::TypedArray => clone_typedarray(obj, id, seen),
        // Per spec: functions, host functions — DataCloneError.
        // MVP returns null.
        KindTag::Function | KindTag::HostFunction => Value::Null,
    }
}

enum KindTag {
    Ordinary, Array, Map, Set, ArrayBuffer, TypedArray, Function, HostFunction,
}

fn kind_discriminant(k: &ObjectKind) -> KindTag {
    match k {
        ObjectKind::Ordinary => KindTag::Ordinary,
        ObjectKind::Array(_) => KindTag::Array,
        ObjectKind::Map(_) => KindTag::Map,
        ObjectKind::Set(_) => KindTag::Set,
        ObjectKind::ArrayBuffer(_) => KindTag::ArrayBuffer,
        ObjectKind::TypedArray(_) => KindTag::TypedArray,
        ObjectKind::Function(_) => KindTag::Function,
        ObjectKind::HostFunction(_) => KindTag::HostFunction,
    }
}

fn clone_ordinary(src: &Arc<Mutex<Object>>, id: usize, seen: &mut HashMap<usize, Value>) -> Value {
    // Allocate empty target first, insert into seen, then copy
    // property values (so cycles resolve to the same target).
    let target_arc = Arc::new(Mutex::new(Object::new()));
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let entries: Vec<(String, Value)> = {
        let s = src.lock().unwrap();
        s.properties.iter()
            .filter(|(k, _)| !k.starts_with("__"))
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect()
    };
    {
        let mut t = target_arc.lock().unwrap();
        for (k, v) in entries {
            t.properties.insert(k, deep_clone(&v, seen));
        }
    }
    target_val
}

fn clone_array(src: &Arc<Mutex<Object>>, id: usize, seen: &mut HashMap<usize, Value>) -> Value {
    let target_arc = Arc::new(Mutex::new(Object::new_array(Vec::new())));
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let elems: Vec<Value> = {
        let s = src.lock().unwrap();
        if let ObjectKind::Array(ref v) = s.kind { v.clone() } else { Vec::new() }
    };
    {
        let mut t = target_arc.lock().unwrap();
        let new_len = {
            if let ObjectKind::Array(ref mut v) = t.kind {
                *v = elems.iter().map(|e| deep_clone(e, seen)).collect();
                v.len()
            } else {
                0
            }
        };
        t.properties.insert("length".into(), Value::F64(new_len as f64));
    }
    target_val
}

fn clone_map(src: &Arc<Mutex<Object>>, id: usize, seen: &mut HashMap<usize, Value>) -> Value {
    let mut target_obj = Object::new();
    target_obj.kind = ObjectKind::Map(indexmap::IndexMap::new());
    target_obj.properties.insert("size".into(), Value::I32(0));
    let target_arc = Arc::new(Mutex::new(target_obj));
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let entries: Vec<(Value, Value)> = {
        let s = src.lock().unwrap();
        if let ObjectKind::Map(ref m) = s.kind {
            m.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
        } else {
            Vec::new()
        }
    };
    {
        let mut t = target_arc.lock().unwrap();
        let new_size = {
            if let ObjectKind::Map(ref mut m) = t.kind {
                for (k, v) in entries {
                    // Per spec: Map keys are cloned too (unlike JSON).
                    m.insert(deep_clone(&k, seen), deep_clone(&v, seen));
                }
                m.len()
            } else {
                0
            }
        };
        t.properties.insert("size".into(), Value::I32(new_size as i32));
    }
    target_val
}

fn clone_set(src: &Arc<Mutex<Object>>, id: usize, seen: &mut HashMap<usize, Value>) -> Value {
    let mut target_obj = Object::new();
    target_obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    target_obj.properties.insert("size".into(), Value::I32(0));
    let target_arc = Arc::new(Mutex::new(target_obj));
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let elements: Vec<Value> = {
        let s = src.lock().unwrap();
        if let ObjectKind::Set(ref set) = s.kind {
            set.iter().cloned().collect()
        } else {
            Vec::new()
        }
    };
    {
        let mut t = target_arc.lock().unwrap();
        let new_size = {
            if let ObjectKind::Set(ref mut set) = t.kind {
                for v in elements {
                    set.insert(deep_clone(&v, seen));
                }
                set.len()
            } else {
                0
            }
        };
        t.properties.insert("size".into(), Value::I32(new_size as i32));
    }
    target_val
}

fn clone_arraybuffer(src: &Arc<Mutex<Object>>, id: usize, seen: &mut HashMap<usize, Value>) -> Value {
    // Per spec: ArrayBuffer clones copy bytes into a fresh buffer.
    let (bytes_copy, max_byte_length, resizable, shared) = {
        let s = src.lock().unwrap();
        if let ObjectKind::ArrayBuffer(ref state) = s.kind {
            let bytes = state.bytes.lock().unwrap().clone();
            (bytes, state.max_byte_length, state.resizable, state.shared)
        } else {
            (Vec::new(), 0, false, false)
        }
    };
    let state = ArrayBufferState {
        bytes: Arc::new(Mutex::new(bytes_copy)),
        max_byte_length,
        resizable,
        detached: false,
        shared,
    };
    let byte_len = state.bytes.lock().unwrap().len();
    let mut obj = Object::new();
    obj.kind = ObjectKind::ArrayBuffer(state);
    obj.properties.insert("byteLength".into(), Value::I32(byte_len as i32));
    obj.properties.insert("maxByteLength".into(), Value::I32(max_byte_length as i32));
    let out = Value::Object(Arc::new(Mutex::new(obj)));
    seen.insert(id, out.clone());
    out
}

fn clone_typedarray(src: &Arc<Mutex<Object>>, id: usize, seen: &mut HashMap<usize, Value>) -> Value {
    // Clone a typed array by cloning its underlying ArrayBuffer and
    // building a fresh view over the copy. The view metadata (elem,
    // byte_offset, length) stays the same.
    let (elem, byte_offset, length, src_bytes) = {
        let s = src.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = s.kind {
            let bytes = ta.buffer.lock().unwrap().clone();
            (ta.elem, ta.byte_offset, ta.length, bytes)
        } else {
            return Value::Null;
        }
    };

    // Fresh ArrayBuffer wrapping the copied bytes.
    let new_bytes = Arc::new(Mutex::new(src_bytes));
    let byte_length = new_bytes.lock().unwrap().len();
    let ab_state = ArrayBufferState {
        bytes: new_bytes.clone(),
        max_byte_length: byte_length,
        resizable: false,
        detached: false,
        shared: false,
    };
    let mut ab_obj = Object::new();
    ab_obj.kind = ObjectKind::ArrayBuffer(ab_state);
    ab_obj.properties.insert("byteLength".into(), Value::I32(byte_length as i32));
    ab_obj.properties.insert("maxByteLength".into(), Value::I32(byte_length as i32));
    let ab_arc = Arc::new(Mutex::new(ab_obj));

    let ta_state = TypedArrayState {
        elem,
        buffer: new_bytes,
        buffer_obj: ab_arc,
        byte_offset,
        length,
    };
    let bpe = elem.bytes_per_element();
    let mut ta_obj = Object::new();
    ta_obj.kind = ObjectKind::TypedArray(ta_state);
    ta_obj.properties.insert("length".into(), Value::I32(length as i32));
    ta_obj.properties.insert("byteLength".into(), Value::I32((length * bpe) as i32));
    ta_obj.properties.insert("byteOffset".into(), Value::I32(byte_offset as i32));
    ta_obj.properties.insert("BYTES_PER_ELEMENT".into(), Value::I32(bpe as i32));
    let out = Value::Object(Arc::new(Mutex::new(ta_obj)));
    seen.insert(id, out.clone());
    out
}
