//! # `ecma:structured-clone` host handler
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

use std::collections::{BTreeSet, HashMap, HashSet};
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{ArrayBufferState, Object, ObjectKind, Value};
use vybe_bytecode::VM;

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:structured-clone",
        "clone",
        Box::new(|ctx, args| {
            if args.is_empty() {
                ctx.throw_value(crate::ecma::error::new_error(
                    ctx,
                    "TypeError",
                    "structuredClone requires at least 1 argument",
                ));
                return Value::Undefined;
            }
            if let Some(error) = validate_transfer_options(ctx, args.get(1)) {
                ctx.throw_value(error);
                return Value::Null;
            }
            let transfer_list = match collect_transfer_list(ctx, args.get(1)) {
                Ok(list) => list,
                Err(error) => {
                    ctx.throw_value(error);
                    return Value::Null;
                }
            };
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let mut seen: HashMap<usize, Value> = HashMap::new();
            let mut active: HashSet<usize> = HashSet::new();
            match deep_clone(ctx, &value, &mut seen, &mut active) {
                Ok(clone) => {
                    detach_transfer_list(&transfer_list);
                    mark_transferred_views(&value, &transfer_list, &mut HashSet::new());
                    clone
                }
                Err(error) => {
                    ctx.throw_value(error);
                    Value::Null
                }
            }
        }),
    );
}

fn deep_clone(
    ctx: &mut vybe_bytecode::HostContext,
    v: &Value,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
) -> Result<Value, Value> {
    match v {
        // Primitives — pass through; they're Copy or immutable.
        Value::Null | Value::TypedNull(_) | Value::Undefined => Ok(v.clone()),
        Value::Bool(_) | Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) => {
            Ok(v.clone())
        }
        Value::BigInt(_) | Value::V128(_) => Ok(v.clone()),

        // Strings: Arc<str> is already cheap to clone; structured
        // clone of a string = the same string per spec.
        Value::String(_) => Ok(v.clone()),

        // Symbols: not cloneable per spec (structured clone throws
        // DataCloneError). MVP passes through — the compiler layer
        // can enforce the throw when we have exception dispatch.
        Value::Symbol(_) => Err(crate::ecma::error::new_error_flat(
            "DataCloneError",
            "symbol could not be cloned",
        )),

        // WeakRef: structured clone of a weak reference yields a
        // dead weak reference per spec. Pass through.
        Value::WeakRef(_) => Ok(v.clone()),

        Value::Object(obj) => clone_object(ctx, obj, seen, active),
    }
}

fn clone_object(
    ctx: &mut vybe_bytecode::HostContext,
    obj: &Arc<Mutex<Object>>,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
) -> Result<Value, Value> {
    let id = Arc::as_ptr(obj) as usize;
    // Cycle handling (HTML structured-clone preserves the reference graph, incl.
    // circular references): every container cloner inserts its freshly-allocated
    // target into `seen` BEFORE recursing, so a back-reference to an in-progress
    // object resolves to that same target — `clone.self === clone`. Check `seen`
    // first so cycles are preserved rather than rejected.
    if let Some(already) = seen.get(&id) {
        return Ok(already.clone());
    }
    if is_uncloneable_weak_collection(obj) {
        return Err(crate::ecma::error::new_error_flat(
            "DataCloneError",
            "weak collections could not be cloned",
        ));
    }

    // Determine the kind first so we can place the freshly-allocated
    // target into `seen` before recursing — needed for cycle handling.
    let kind_tag = {
        let o = obj.lock().unwrap();
        kind_discriminant(&o.kind)
    };

    active.insert(id);
    let result = match kind_tag {
        KindTag::Ordinary => clone_ordinary(ctx, obj, id, seen, active),
        KindTag::Array => clone_array(ctx, obj, id, seen, active),
        KindTag::Map => clone_map(ctx, obj, id, seen, active),
        KindTag::Set => clone_set(ctx, obj, id, seen, active),
        KindTag::ArrayBuffer => clone_arraybuffer(obj, id, seen),
        KindTag::TypedArray => clone_typedarray(obj, id, seen),
        // HTML structured-clone (§ StructuredSerializeInternal): a callable
        // (function / host function / generator continuation) is not
        // serializable → throw a DataCloneError. `?`-propagated by the array /
        // object / map / set cloners, so a function nested anywhere aborts the
        // whole clone, matching the spec and browsers.
        KindTag::Function | KindTag::HostFunction | KindTag::Continuation => {
            Err(crate::ecma::error::new_error_flat(
                "DataCloneError",
                "could not be cloned (functions are not structured-cloneable)",
            ))
        }
        // Module Namespace Objects are frozen spec-exotic objects;
        // structuredClone on them is not spec-defined. Futures/streams are
        // VM-internal. `null` is the conservative result for these.
        KindTag::ModuleNamespace | KindTag::Future | KindTag::Stream => Ok(Value::Null),
    };
    active.remove(&id);
    result
}

enum KindTag {
    Ordinary,
    Array,
    Map,
    Set,
    ArrayBuffer,
    TypedArray,
    Function,
    HostFunction,
    Continuation,
    ModuleNamespace,
    Future,
    Stream,
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
        ObjectKind::Continuation(_) => KindTag::Continuation,
        ObjectKind::ModuleNamespace => KindTag::ModuleNamespace,
        ObjectKind::Future { .. } => KindTag::Future,
        ObjectKind::Stream { .. } => KindTag::Stream,
    }
}

fn clone_ordinary(
    ctx: &mut vybe_bytecode::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
) -> Result<Value, Value> {
    if let Some((target, _handler)) = crate::ecma::object::proxy_target_and_handler(src) {
        return match target {
            Value::Object(target_obj) => clone_object(ctx, &target_obj, seen, active),
            other => deep_clone(ctx, &other, seen, active),
        };
    }
    if let Some(cloned) = clone_builtin_object(src, id, seen) {
        return cloned;
    }
    if let Some(error_kind) = error_clone_kind(src) {
        return clone_error_like(ctx, src, id, seen, active, &error_kind);
    }
    if is_dataview_object(src) {
        return clone_dataview(src, id, seen);
    }

    // Allocate empty target first, insert into seen, then copy
    // property values (so cycles resolve to the same target).
    let target_arc = vybe_bytecode::heap::alloc(Object::new());
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let entries: Vec<(String, Value)> = {
        let s = src.lock().unwrap();
        s.properties
            .iter()
            .filter(|(k, _)| {
                !k.starts_with("__")
                    && !k.starts_with("Symbol(")
                    && !crate::ecma::object::is_nonenum(&s, k)
            })
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect()
    };
    let accessors: Vec<(String, Value)> = {
        let s = src.lock().unwrap();
        s.properties
            .iter()
            .filter_map(|(k, getter)| {
                let prop = k.strip_prefix("__get_")?;
                if prop.starts_with("__")
                    || prop.starts_with("Symbol(")
                    || crate::ecma::object::is_nonenum(&s, prop)
                {
                    return None;
                }
                Some((prop.to_string(), getter.clone()))
            })
            .collect()
    };
    {
        let mut t = target_arc.lock().unwrap();
        for (k, v) in entries {
            t.properties.insert(k, deep_clone(ctx, &v, seen, active)?);
        }
    }
    for (k, getter) in accessors {
        let value = ctx.invoke(&getter, &[Value::Object(src.clone())]);
        let cloned = deep_clone(ctx, &value, seen, active)?;
        target_arc.lock().unwrap().properties.insert(k, cloned);
    }
    Ok(target_val)
}

fn clone_builtin_object(
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Option<Result<Value, Value>> {
    let (type_tag, primitive, time, regexp_fields) = {
        let s = src.lock().unwrap();
        let type_tag = string_prop(&s, "__type");
        let primitive = s.properties.get("__primitive").cloned();
        let time = s.properties.get("__time").cloned();
        let regexp_fields = if matches!(type_tag.as_deref(), Some("RegExp")) {
            Some((
                s.properties
                    .get("source")
                    .cloned()
                    .unwrap_or_else(|| Value::String(Arc::from("(?:)"))),
                s.properties
                    .get("flags")
                    .cloned()
                    .unwrap_or_else(|| Value::String(Arc::from(""))),
                s.properties
                    .get("lastIndex")
                    .cloned()
                    .unwrap_or(Value::I32(0)),
            ))
        } else {
            None
        };
        (type_tag, primitive, time, regexp_fields)
    };

    let out = match type_tag.as_deref()? {
        "Date" => {
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Date")));
            obj.properties
                .insert("__time".into(), time.unwrap_or(Value::F64(f64::NAN)));
            obj.properties.insert(
                "__proto__".into(),
                crate::ecma::date::shared_date_prototype(),
            );
            Value::Object(vybe_bytecode::heap::alloc(obj))
        }
        "RegExp" => {
            let (source, flags, last_index) = regexp_fields?;
            let source_text = match source {
                Value::String(s) => s.to_string(),
                other => format!("{}", other),
            };
            let flags_text = match flags {
                Value::String(s) => s.to_string(),
                other => format!("{}", other),
            };
            let mut obj = Object::new();
            obj.properties.insert(
                "source".into(),
                Value::String(Arc::from(source_text.as_str())),
            );
            obj.properties.insert(
                "flags".into(),
                Value::String(Arc::from(flags_text.as_str())),
            );
            obj.properties
                .insert("global".into(), Value::Bool(flags_text.contains('g')));
            obj.properties
                .insert("ignoreCase".into(), Value::Bool(flags_text.contains('i')));
            obj.properties
                .insert("multiline".into(), Value::Bool(flags_text.contains('m')));
            obj.properties
                .insert("dotAll".into(), Value::Bool(flags_text.contains('s')));
            obj.properties
                .insert("unicode".into(), Value::Bool(flags_text.contains('u')));
            obj.properties
                .insert("unicodeSets".into(), Value::Bool(flags_text.contains('v')));
            obj.properties
                .insert("sticky".into(), Value::Bool(flags_text.contains('y')));
            obj.properties
                .insert("hasIndices".into(), Value::Bool(flags_text.contains('d')));
            obj.properties.insert("lastIndex".into(), last_index);
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("RegExp")));
            obj.properties.insert(
                "__proto__".into(),
                crate::ecma::regexp::shared_regexp_prototype(),
            );
            Value::Object(vybe_bytecode::heap::alloc(obj))
        }
        "Boolean" => match primitive {
            Some(Value::Bool(value)) => crate::ecma::boolean::boxed_boolean(value),
            Some(value) => {
                crate::ecma::boolean::boxed_boolean(crate::ecma::boolean::to_boolean(&value))
            }
            None => crate::ecma::boolean::boxed_boolean(false),
        },
        "Number" => crate::ecma::number::boxed_number(primitive.unwrap_or(Value::F64(f64::NAN))),
        "String" => {
            let text = match primitive {
                Some(Value::String(s)) => s,
                Some(value) => Arc::from(format!("{}", value).as_str()),
                None => Arc::from(""),
            };
            crate::ecma::string::boxed_string(text)
        }
        _ => return None,
    };
    seen.insert(id, out.clone());
    Some(Ok(out))
}

fn clone_array(
    ctx: &mut vybe_bytecode::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
) -> Result<Value, Value> {
    let target_arc = vybe_bytecode::heap::alloc(Object::new_array(Vec::new()));
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let (elems, holes): (Vec<Value>, Vec<usize>) = {
        let s = src.lock().unwrap();
        if let ObjectKind::Array(ref v) = s.kind {
            let holes = (0..v.len())
                .filter(|index| crate::ecma::array::is_array_hole(&s, *index))
                .collect();
            (v.clone(), holes)
        } else {
            (Vec::new(), Vec::new())
        }
    };
    {
        let mut t = target_arc.lock().unwrap();
        let cloned_elems: Result<Vec<Value>, Value> = elems
            .iter()
            .map(|e| deep_clone(ctx, e, seen, active))
            .collect();
        let new_len = {
            if let ObjectKind::Array(ref mut v) = t.kind {
                *v = cloned_elems?;
                v.len()
            } else {
                0
            }
        };
        t.properties
            .insert("length".into(), Value::F64(new_len as f64));
        let hole_set: BTreeSet<usize> = holes.into_iter().collect();
        crate::ecma::array::store_hole_indices(&mut t, &hole_set);
    }
    Ok(target_val)
}

fn clone_map(
    ctx: &mut vybe_bytecode::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
) -> Result<Value, Value> {
    let mut target_obj = Object::new();
    target_obj.kind = ObjectKind::Map(indexmap::IndexMap::new());
    target_obj.properties.insert("size".into(), Value::I32(0));
    target_obj
        .properties
        .insert("__type".into(), Value::String(Arc::from("Map")));
    let target_arc = vybe_bytecode::heap::alloc(target_obj);
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
                    m.insert(
                        deep_clone(ctx, &k, seen, active)?,
                        deep_clone(ctx, &v, seen, active)?,
                    );
                }
                m.len()
            } else {
                0
            }
        };
        t.properties
            .insert("size".into(), Value::I32(new_size as i32));
    }
    Ok(target_val)
}

fn clone_set(
    ctx: &mut vybe_bytecode::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
) -> Result<Value, Value> {
    let mut target_obj = Object::new();
    target_obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    target_obj.properties.insert("size".into(), Value::I32(0));
    target_obj
        .properties
        .insert("__type".into(), Value::String(Arc::from("Set")));
    let target_arc = vybe_bytecode::heap::alloc(target_obj);
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
                    set.insert(deep_clone(ctx, &v, seen, active)?);
                }
                set.len()
            } else {
                0
            }
        };
        t.properties
            .insert("size".into(), Value::I32(new_size as i32));
    }
    Ok(target_val)
}

fn clone_arraybuffer(
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    // Per spec: ArrayBuffer clones copy bytes into a fresh buffer; SharedArrayBuffer
    // stays shared and is not duplicated.
    let (bytes_copy, max_byte_length, resizable, shared, detached) = {
        let s = src.lock().unwrap();
        if let ObjectKind::ArrayBuffer(ref state) = s.kind {
            let bytes = state.bytes.lock().unwrap().clone();
            (
                bytes,
                state.max_byte_length,
                state.resizable,
                state.shared,
                state.detached,
            )
        } else {
            (Vec::new(), 0, false, false, false)
        }
    };
    if detached {
        return Err(crate::ecma::error::new_error_flat(
            "DataCloneError",
            "ArrayBuffer is detached",
        ));
    }
    if shared {
        let out = Value::Object(src.clone());
        seen.insert(id, out.clone());
        return Ok(out);
    }
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
    obj.properties
        .insert("byteLength".into(), Value::I32(byte_len as i32));
    obj.properties
        .insert("maxByteLength".into(), Value::I32(max_byte_length as i32));
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("ArrayBuffer")));
    let out = Value::Object(vybe_bytecode::heap::alloc(obj));
    seen.insert(id, out.clone());
    Ok(out)
}

fn clone_typedarray(
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    let (elem, byte_offset, length, buffer_obj, src_bytes, buffer_detached) = {
        let s = src.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = s.kind {
            let bytes = ta.buffer.lock().unwrap().clone();
            let detached = ta.buffer_obj.lock().unwrap().kind.as_arraybuffer_detached();
            (
                ta.elem,
                ta.byte_offset,
                ta.length,
                ta.buffer_obj.clone(),
                bytes,
                detached,
            )
        } else {
            return Ok(Value::Null);
        }
    };
    if buffer_detached {
        return Err(crate::ecma::error::new_error_flat(
            "DataCloneError",
            "ArrayBuffer is detached",
        ));
    }

    let bpe = elem.bytes_per_element();
    let view_bytes = length.saturating_mul(bpe);
    let spans_entire_buffer = byte_offset == 0 && view_bytes == src_bytes.len();
    let out = if spans_entire_buffer {
        let buffer_id = Arc::as_ptr(&buffer_obj) as usize;
        let cloned_buffer = if let Some(Value::Object(existing)) = seen.get(&buffer_id) {
            existing.clone()
        } else {
            match clone_arraybuffer(&buffer_obj, buffer_id, seen)? {
                Value::Object(obj) => obj,
                _ => return Ok(Value::Null),
            }
        };
        crate::ecma::typedarray::new_view_over_buffer(elem, cloned_buffer, 0, length)
    } else {
        let out = crate::ecma::typedarray::new_typed_array(elem, length);
        if let Value::Object(ref target) = out {
            let mut t = target.lock().unwrap();
            if let ObjectKind::TypedArray(ref mut ta) = t.kind {
                for i in 0..length {
                    let src_index = byte_offset + i * bpe;
                    let end = src_index + bpe;
                    let mut bytes = ta.buffer.lock().unwrap();
                    if end <= src_bytes.len() && i * bpe + bpe <= bytes.len() {
                        bytes[i * bpe..i * bpe + bpe].copy_from_slice(&src_bytes[src_index..end]);
                    }
                }
            }
        }
        out
    };
    seen.insert(id, out.clone());
    Ok(out)
}

trait ArrayBufferDetached {
    fn as_arraybuffer_detached(&self) -> bool;
}

impl ArrayBufferDetached for ObjectKind {
    fn as_arraybuffer_detached(&self) -> bool {
        matches!(self, ObjectKind::ArrayBuffer(state) if state.detached)
    }
}

fn is_dataview_object(src: &Arc<Mutex<Object>>) -> bool {
    let s = src.lock().unwrap();
    s.properties
        .get(crate::ecma::arraybuffer::DV_TAG)
        .is_some_and(|v| !matches!(v, Value::Undefined | Value::Null | Value::Bool(false)))
}

fn is_uncloneable_weak_collection(src: &Arc<Mutex<Object>>) -> bool {
    let s = src.lock().unwrap();
    matches!(
        s.properties.get("__type"),
        Some(Value::String(tag)) if matches!(tag.as_ref(), "WeakMap" | "WeakSet")
    ) || s.properties.contains_key(crate::ecma::weakmap::WEAKMAP_TAG)
        || s.properties.contains_key(crate::ecma::weakmap::WEAKSET_TAG)
}

fn clone_dataview(
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    let (buffer, byte_offset, byte_length) = {
        let s = src.lock().unwrap();
        let buffer = match s.properties.get("buffer").cloned() {
            Some(Value::Object(buffer)) => buffer,
            _ => return Ok(Value::Null),
        };
        let byte_offset = match s.properties.get("byteOffset") {
            Some(Value::I32(n)) => (*n).max(0) as usize,
            Some(v) => v.as_i32().max(0) as usize,
            _ => 0,
        };
        let byte_length = match s.properties.get("byteLength") {
            Some(Value::I32(n)) => (*n).max(0) as usize,
            Some(v) => v.as_i32().max(0) as usize,
            _ => 0,
        };
        (buffer, byte_offset, byte_length)
    };
    let bytes = {
        let b = buffer.lock().unwrap();
        if let ObjectKind::ArrayBuffer(ref state) = b.kind {
            if state.detached {
                return Err(crate::ecma::error::new_error_flat(
                    "DataCloneError",
                    "ArrayBuffer is detached",
                ));
            }
            let src_bytes = state.bytes.lock().unwrap();
            let end = byte_offset.saturating_add(byte_length).min(src_bytes.len());
            if byte_offset < end {
                src_bytes[byte_offset..end].to_vec()
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        }
    };
    let byte_length = bytes.len();
    let state = ArrayBufferState {
        bytes: Arc::new(Mutex::new(bytes)),
        max_byte_length: byte_length,
        resizable: false,
        detached: false,
        shared: false,
    };
    let mut obj = Object::new();
    obj.kind = ObjectKind::ArrayBuffer(state);
    obj.properties
        .insert("byteLength".into(), Value::I32(byte_length as i32));
    obj.properties
        .insert("maxByteLength".into(), Value::I32(byte_length as i32));
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("ArrayBuffer")));
    let buffer_value = Value::Object(vybe_bytecode::heap::alloc(obj));
    let out = crate::ecma::arraybuffer::new_dataview(buffer_value, 0, byte_length as i32);
    seen.insert(id, out.clone());
    Ok(out)
}

fn standard_error_kind(kind: &str) -> bool {
    matches!(
        kind,
        "Error"
            | "TypeError"
            | "RangeError"
            | "SyntaxError"
            | "ReferenceError"
            | "URIError"
            | "EvalError"
            | "AggregateError"
            | "SuppressedError"
    )
}

fn error_clone_kind(src: &Arc<Mutex<Object>>) -> Option<String> {
    let s = src.lock().unwrap();
    [
        string_prop(&s, "__exception_type"),
        string_prop(&s, "__type"),
        string_prop(&s, "name"),
        proto_string_prop(&s, "name"),
        error_kind_from_tag(&s),
    ]
    .into_iter()
    .flatten()
    .find(|kind| standard_error_kind(kind) || kind.ends_with("Error"))
}

fn clone_error_like(
    ctx: &mut vybe_bytecode::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
    active: &mut HashSet<usize>,
    source_kind: &str,
) -> Result<Value, Value> {
    let target_arc = vybe_bytecode::heap::alloc(Object::new());
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let (entries, display_name, message) = {
        let s = src.lock().unwrap();
        let display_name = string_prop(&s, "name")
            .or_else(|| proto_string_prop(&s, "name"))
            .unwrap_or_else(|| source_kind.to_string());
        let message = string_prop(&s, "message").unwrap_or_default();
        let entries: Vec<(String, Value)> = s
            .properties
            .iter()
            .filter(|(k, _)| should_clone_error_property(&s, k))
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect();
        (entries, display_name, message)
    };

    let clone_kind = if standard_error_kind(source_kind) {
        source_kind
    } else {
        "Error"
    };
    {
        let mut t = target_arc.lock().unwrap();
        stamp_error_clone(&mut t, clone_kind, &display_name, &message);
    }
    for (k, v) in entries {
        let cloned = deep_clone(ctx, &v, seen, active)?;
        target_arc.lock().unwrap().properties.insert(k, cloned);
    }
    Ok(target_val)
}

fn should_clone_error_property(o: &Object, key: &str) -> bool {
    if key.starts_with("__")
        || key.starts_with("Symbol(")
        || key.starts_with("__get_")
        || key.starts_with("__set_")
    {
        return false;
    }
    matches!(
        key,
        "name" | "message" | "stack" | "cause" | "errors" | "error" | "suppressed"
    ) || !crate::ecma::object::is_nonenum(o, key)
}

fn stamp_error_clone(obj: &mut Object, kind: &str, display_name: &str, message: &str) {
    obj.properties
        .insert("__type".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("__exception_type".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("name".into(), Value::String(Arc::from(display_name)));
    obj.properties
        .insert("message".into(), Value::String(Arc::from(message)));
    obj.properties.insert(
        "stack".into(),
        Value::String(Arc::from(format!("{}: {}", display_name, message).as_str())),
    );
    let chain: Vec<Value> = error_ancestors(kind)
        .iter()
        .map(|n| Value::String(Arc::from(*n)))
        .collect();
    obj.properties.insert(
        "__types".into(),
        Value::Object(vybe_bytecode::heap::alloc(Object::new_array(chain))),
    );
}

fn error_ancestors(kind: &str) -> &'static [&'static str] {
    match kind {
        "Error" => &["Error"],
        "TypeError" => &["TypeError", "Error"],
        "RangeError" => &["RangeError", "Error"],
        "SyntaxError" => &["SyntaxError", "Error"],
        "ReferenceError" => &["ReferenceError", "Error"],
        "URIError" => &["URIError", "Error"],
        "EvalError" => &["EvalError", "Error"],
        "AggregateError" => &["AggregateError", "Error"],
        "SuppressedError" => &["SuppressedError", "Error"],
        _ => &["Error"],
    }
}

fn string_prop(o: &Object, key: &str) -> Option<String> {
    match o.properties.get(key) {
        Some(Value::String(s)) => Some(s.to_string()),
        Some(v) => Some(format!("{}", v)),
        None => None,
    }
}

fn proto_string_prop(o: &Object, key: &str) -> Option<String> {
    let mut current = match o.properties.get("__proto__") {
        Some(Value::Object(proto)) => Some(proto.clone()),
        _ => None,
    };
    let mut seen = HashSet::new();
    while let Some(obj) = current {
        let id = Arc::as_ptr(&obj) as usize;
        if !seen.insert(id) {
            return None;
        }
        let guard = obj.lock().unwrap();
        if let Some(value) = string_prop(&guard, key) {
            return Some(value);
        }
        current = match guard.properties.get("__proto__") {
            Some(Value::Object(next)) => Some(next.clone()),
            _ => None,
        };
    }
    None
}

fn error_kind_from_tag(o: &Object) -> Option<String> {
    let tag = string_prop(o, "tostringtag")?;
    let tag = tag
        .strip_prefix("[object ")
        .and_then(|s| s.strip_suffix(']'))
        .unwrap_or(tag.as_str());
    if standard_error_kind(tag) || tag.ends_with("Error") {
        Some(tag.to_string())
    } else {
        None
    }
}

fn validate_transfer_options(
    _ctx: &vybe_bytecode::HostContext,
    options: Option<&Value>,
) -> Option<Value> {
    match options {
        None | Some(Value::Undefined) => None,
        Some(Value::Object(_)) => None,
        Some(_) => Some(crate::ecma::error::new_error_flat(
            "TypeError",
            "structuredClone options must be an object",
        )),
    }
}

fn collect_transfer_list(
    ctx: &mut vybe_bytecode::HostContext,
    options: Option<&Value>,
) -> Result<Vec<Arc<Mutex<Object>>>, Value> {
    let Some(Value::Object(options)) = options else {
        return Ok(Vec::new());
    };
    let transfer = {
        let o = options.lock().unwrap();
        if let Some(getter) = o.properties.get("__get_transfer").cloned() {
            drop(o);
            let arity = match &getter {
                Value::Object(getter_obj) => match &getter_obj.lock().unwrap().kind {
                    ObjectKind::Function(function) => function.arity,
                    _ => 0,
                },
                _ => 0,
            };
            let receiver = Value::Object(options.clone());
            let saved_this = ctx.current_js_this();
            ctx.set_js_this(receiver.clone());
            let result = if arity >= 1 {
                ctx.try_invoke(&getter, &[receiver])
            } else {
                ctx.try_invoke(&getter, &[])
            };
            ctx.set_js_this(saved_this);
            Some(result?)
        } else {
            o.properties.get("transfer").cloned()
        }
    };
    let Some(transfer) = transfer else {
        return Ok(Vec::new());
    };
    if matches!(transfer, Value::Undefined) {
        return Ok(Vec::new());
    }
    let mut out = Vec::new();
    match transfer {
        Value::Object(list) => {
            let l = list.lock().unwrap();
            match &l.kind {
                ObjectKind::Array(values) => {
                    for value in values {
                        push_transferable(value, &mut out)?;
                    }
                }
                ObjectKind::Set(values) => {
                    for value in values {
                        push_transferable(value, &mut out)?;
                    }
                }
                _ => {
                    return Err(crate::ecma::error::new_error_flat(
                        "TypeError",
                        "transfer option is not iterable",
                    ));
                }
            }
        }
        Value::Null => {
            return Err(crate::ecma::error::new_error_flat(
                "TypeError",
                "transfer list contains null",
            ));
        }
        _ => {
            return Err(crate::ecma::error::new_error_flat(
                "TypeError",
                "transfer option is not iterable",
            ));
        }
    }
    let mut ids = HashSet::new();
    for item in &out {
        let id = Arc::as_ptr(item) as usize;
        if !ids.insert(id) {
            return Err(crate::ecma::error::new_error_flat(
                "DataCloneError",
                "duplicate transfer item",
            ));
        }
    }
    Ok(out)
}

fn push_transferable(value: &Value, out: &mut Vec<Arc<Mutex<Object>>>) -> Result<(), Value> {
    let Value::Object(obj) = value else {
        return Err(crate::ecma::error::new_error_flat(
            "DataCloneError",
            "transfer item is not transferable",
        ));
    };
    {
        let o = obj.lock().unwrap();
        let ObjectKind::ArrayBuffer(state) = &o.kind else {
            return Err(crate::ecma::error::new_error_flat(
                "DataCloneError",
                "transfer item is not transferable",
            ));
        };
        if state.shared {
            return Err(crate::ecma::error::new_error_flat(
                "DataCloneError",
                "SharedArrayBuffer is not transferable",
            ));
        }
        if state.detached {
            return Err(crate::ecma::error::new_error_flat(
                "DataCloneError",
                "ArrayBuffer is detached",
            ));
        }
    }
    out.push(obj.clone());
    Ok(())
}

fn detach_transfer_list(items: &[Arc<Mutex<Object>>]) {
    for item in items {
        let mut o = item.lock().unwrap();
        if let ObjectKind::ArrayBuffer(ref mut state) = o.kind {
            state.bytes.lock().unwrap().clear();
            state.detached = true;
            state.resizable = false;
            state.max_byte_length = 0;
            state.shared = false;
            o.properties.insert("byteLength".into(), Value::I32(0));
            o.properties.insert("maxByteLength".into(), Value::I32(0));
        }
    }
}

fn mark_transferred_views(
    value: &Value,
    transferred: &[Arc<Mutex<Object>>],
    seen: &mut HashSet<usize>,
) {
    let transferred_ids: HashSet<usize> = transferred
        .iter()
        .map(|item| Arc::as_ptr(item) as usize)
        .collect();
    mark_transferred_views_inner(value, &transferred_ids, seen);
}

fn mark_transferred_views_inner(
    value: &Value,
    transferred_ids: &HashSet<usize>,
    seen: &mut HashSet<usize>,
) {
    let Value::Object(obj) = value else {
        return;
    };
    let id = Arc::as_ptr(obj) as usize;
    if !seen.insert(id) {
        return;
    }
    let children: Vec<Value> = {
        let mut o = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            let buffer_id = Arc::as_ptr(&ta.buffer_obj) as usize;
            if transferred_ids.contains(&buffer_id) {
                o.properties.insert("length".into(), Value::I32(0));
                o.properties.insert("byteLength".into(), Value::I32(0));
            }
        } else if o.properties.contains_key(crate::ecma::arraybuffer::DV_TAG) {
            let buffer_id = match o.properties.get("buffer") {
                Some(Value::Object(buffer)) => Arc::as_ptr(buffer) as usize,
                _ => 0,
            };
            if transferred_ids.contains(&buffer_id) {
                o.properties.insert("byteLength".into(), Value::I32(0));
            }
        }
        let mut values = Vec::new();
        match &o.kind {
            ObjectKind::Array(items) => values.extend(items.iter().cloned()),
            ObjectKind::Map(map) => {
                for (k, v) in map {
                    values.push(k.clone());
                    values.push(v.clone());
                }
            }
            ObjectKind::Set(set) => values.extend(set.iter().cloned()),
            _ => values.extend(
                o.properties
                    .iter()
                    .filter(|(key, _)| !key.starts_with("__"))
                    .map(|(_, value)| value.clone()),
            ),
        }
        values
    };
    for child in children {
        mark_transferred_views_inner(&child, transferred_ids, seen);
    }
}
