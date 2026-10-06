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
use vybe_runtime::VM;
use vybe_runtime::value::{ArrayBufferState, Object, ObjectKind, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:structured-clone",
        "clone",
        Box::new(|ctx, args| {
            if args.is_empty() {
                ctx.throw_value(crate::error::new_error(
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
            let value = &args[0];
            let mut seen: HashMap<usize, Value> = HashMap::new();
            match deep_clone(ctx, value, &mut seen) {
                Ok(clone) => {
                    detach_transfer_list(&transfer_list);
                    mark_transferred_views(value, &transfer_list, &mut HashSet::new());
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
    ctx: &mut vybe_runtime::HostContext,
    v: &Value,
    seen: &mut HashMap<usize, Value>,
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
        Value::Symbol(_) => Err(crate::error::new_error_flat(
            "DataCloneError",
            "symbol could not be cloned",
        )),

        // WeakRef: structured clone of a weak reference yields a
        // dead weak reference per spec. Pass through.
        Value::WeakRef(_) => Ok(v.clone()),

        Value::Object(obj) => clone_object(ctx, obj, seen),
    }
}

fn clone_object(
    ctx: &mut vybe_runtime::HostContext,
    obj: &Arc<Mutex<Object>>,
    seen: &mut HashMap<usize, Value>,
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
    // Check unsupported brands and classify under one source lock. The
    // existing seen map alone preserves cycles and repeated references.
    let kind_tag = {
        let o = obj.lock().unwrap();
        if is_uncloneable_weak_collection(&o) {
            return Err(crate::error::new_error_flat(
                "DataCloneError",
                "weak collections could not be cloned",
            ));
        }
        kind_discriminant(&o.kind)
    };

    match kind_tag {
        KindTag::Ordinary => clone_ordinary(ctx, obj, id, seen),
        KindTag::Array => clone_array(ctx, obj, id, seen),
        KindTag::Map => clone_map(ctx, obj, id, seen),
        KindTag::Set => clone_set(ctx, obj, id, seen),
        KindTag::ArrayBuffer => clone_arraybuffer(obj, id, seen),
        KindTag::TypedArray => clone_typedarray(obj, id, seen),
        // HTML structured-clone (§ StructuredSerializeInternal): a callable
        // (function / host function / generator continuation) is not
        // serializable → throw a DataCloneError. `?`-propagated by the array /
        // object / map / set cloners, so a function nested anywhere aborts the
        // whole clone, matching the spec and browsers.
        KindTag::Function | KindTag::HostFunction | KindTag::Continuation => {
            Err(crate::error::new_error_flat(
                "DataCloneError",
                "could not be cloned (functions are not structured-cloneable)",
            ))
        }
        // Module Namespace Objects are frozen spec-exotic objects;
        // structuredClone on them is not spec-defined. Futures/streams are
        // VM-internal. `null` is the conservative result for these.
        KindTag::ModuleNamespace | KindTag::Future | KindTag::Stream => Ok(Value::Null),
    }
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

fn ordinary_property_snapshots(source: &Object) -> (Vec<(String, Value)>, Vec<(String, Value)>) {
    let mut entries = Vec::new();
    let mut accessors = Vec::new();
    for (key, value) in &source.properties {
        if let Some(property) = key.strip_prefix("__get_") {
            if !property.starts_with("__")
                && !property.starts_with("Symbol(")
                && !crate::object::is_nonenum(source, property)
            {
                accessors.push((property.to_string(), value.clone()));
            }
        } else if !key.starts_with("__")
            && !key.starts_with("Symbol(")
            && !crate::object::is_nonenum(source, key)
        {
            entries.push((key.clone(), value.clone()));
        }
    }
    (entries, accessors)
}

fn clone_ordinary(
    ctx: &mut vybe_runtime::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    if let Some((target, _handler)) = crate::object::proxy_target_and_handler(src) {
        return match target {
            Value::Object(target_obj) => clone_object(ctx, &target_obj, seen),
            other => deep_clone(ctx, &other, seen),
        };
    }
    if let Some(cloned) = clone_builtin_object(src, id, seen) {
        return cloned;
    }
    if let Some(error_kind) = error_clone_kind(src) {
        return clone_error_like(ctx, src, id, seen, &error_kind);
    }
    if is_dataview_object(src) {
        return clone_dataview(src, id, seen);
    }

    // Allocate empty target first, insert into seen, then copy
    // property values (so cycles resolve to the same target).
    let target_arc = vybe_runtime::heap::alloc(Object::new());
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let (entries, accessors) = {
        let s = src.lock().unwrap();
        ordinary_property_snapshots(&s)
    };
    {
        let mut t = target_arc.lock().unwrap();
        let capacity = entries.len().saturating_add(accessors.len());
        if capacity != 0 {
            t.properties.reserve(capacity);
        }
        for (k, v) in entries {
            t.properties.insert(k, deep_clone(ctx, &v, seen)?);
        }
    }
    for (k, getter) in accessors {
        let value = ctx.invoke(&getter, &[Value::Object(src.clone())]);
        let cloned = deep_clone(ctx, &value, seen)?;
        target_arc.lock().unwrap().properties.insert(k, cloned);
    }
    Ok(target_val)
}

#[cfg(test)]
mod ordinary_complete_clone_benchmark_tests {
    use super::*;
fn previous_clone_ordinary(
    ctx: &mut vybe_runtime::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    if let Some((target, _handler)) = crate::object::proxy_target_and_handler(src) {
        return match target {
            Value::Object(target_obj) => clone_object(ctx, &target_obj, seen),
            other => deep_clone(ctx, &other, seen),
        };
    }
    if let Some(cloned) = clone_builtin_object(src, id, seen) {
        return cloned;
    }
    if let Some(error_kind) = error_clone_kind(src) {
        return clone_error_like(ctx, src, id, seen, &error_kind);
    }
    if is_dataview_object(src) {
        return clone_dataview(src, id, seen);
    }

    // Allocate empty target first, insert into seen, then copy
    // property values (so cycles resolve to the same target).
    let target_arc = vybe_runtime::heap::alloc(Object::new());
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let entries: Vec<(String, Value)> = {
        let source = src.lock().unwrap();
        source.properties.iter().filter(|(key, _)| {
            !key.starts_with("__") && !key.starts_with("Symbol(")
                && !crate::object::is_nonenum(&source, key)
        }).map(|(key, value)| (key.clone(), value.clone())).collect()
    };
    let accessors: Vec<(String, Value)> = {
        let source = src.lock().unwrap();
        source.properties.iter().filter_map(|(key, getter)| {
            let property = key.strip_prefix("__get_")?;
            if property.starts_with("__") || property.starts_with("Symbol(")
                || crate::object::is_nonenum(&source, property)
            {
                return None;
            }
            Some((property.to_string(), getter.clone()))
        }).collect()
    };
    {
        let mut t = target_arc.lock().unwrap();
        for (k, v) in entries {
            t.properties.insert(k, deep_clone(ctx, &v, seen)?);
        }
    }
    for (k, getter) in accessors {
        let value = ctx.invoke(&getter, &[Value::Object(src.clone())]);
        let cloned = deep_clone(ctx, &value, seen)?;
        target_arc.lock().unwrap().properties.insert(k, cloned);
    }
    Ok(target_val)
}


    #[test]
    #[ignore = "paired complete flat ordinary clone benchmark, excludes registered dispatch and nested graphs"]
    fn paired_complete_flat_clone_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        for length in [0, 8, 64, 512] {
            let mut source = Object::new();
            for index in 0..length {
                source.properties.insert(format!("field-{index}"), Value::I32(index));
            }
            let source = vybe_runtime::heap::alloc(source);
            let id = Arc::as_ptr(&source) as usize;
            let iterations = if length > 64 { 100 } else { 1000 };
            let mut previous = Duration::ZERO;
            let mut optimized = Duration::ZERO;
            let mut ctx = vybe_runtime::HostContext::empty();
            for round in 0..8 {
                for current in if round % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..iterations {
                        let mut seen = HashMap::new();
                        let result = if current {
                            clone_ordinary(&mut ctx, black_box(&source), id, &mut seen)
                        } else {
                            previous_clone_ordinary(&mut ctx, black_box(&source), id, &mut seen)
                        }.unwrap();
                        let Value::Object(object) = &result else { panic!("expected cloned object") };
                        assert_eq!(object.lock().unwrap().properties.len(), length as usize);
                        black_box(result);
                    }
                    if current { optimized += start.elapsed(); } else { previous += start.elapsed(); }
                }
            }
            eprintln!("complete flat ordinary clone len={length}: previous={previous:?}, optimized={optimized:?}");
        }
    }
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
                    .unwrap_or_else(|| crate::keys::string_value("(?:)")),
                s.properties
                    .get("flags")
                    .cloned()
                    .unwrap_or_else(|| crate::keys::string_value("")),
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
                .insert("__type".into(), crate::keys::string_value("Date"));
            obj.properties
                .insert("__time".into(), time.unwrap_or(Value::F64(f64::NAN)));
            obj.properties
                .insert("__proto__".into(), crate::date::shared_date_prototype());
            Value::Object(vybe_runtime::heap::alloc(obj))
        }
        "RegExp" => {
            let (source, flags, last_index) = regexp_fields?;
            let source_text = match source {
                Value::String(s) => s.to_string(),
                other => crate::keys::value_display_string(&other),
            };
            let flags_text = match flags {
                Value::String(s) => s.to_string(),
                other => crate::keys::value_display_string(&other),
            };
            let mut obj = Object::new();
            obj.properties
                .insert("source".into(), crate::keys::string_value(&source_text));
            obj.properties
                .insert("flags".into(), crate::keys::string_value(&flags_text));
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
                .insert("__type".into(), crate::keys::string_value("RegExp"));
            obj.properties
                .insert("__proto__".into(), crate::regexp::shared_regexp_prototype());
            Value::Object(vybe_runtime::heap::alloc(obj))
        }
        "Boolean" => match primitive {
            Some(Value::Bool(value)) => crate::boolean::boxed_boolean(value),
            Some(value) => crate::boolean::boxed_boolean(crate::boolean::to_boolean(&value)),
            None => crate::boolean::boxed_boolean(false),
        },
        "Number" => crate::number::boxed_number(primitive.unwrap_or(Value::F64(f64::NAN))),
        "String" => {
            let text = match primitive {
                Some(Value::String(s)) => s,
                Some(value) => {
                    crate::keys::string_arc(crate::keys::value_display_cow(&value).as_ref())
                }
                None => crate::keys::string_arc(""),
            };
            crate::string::boxed_string(text)
        }
        _ => return None,
    };
    seen.insert(id, out.clone());
    Some(Ok(out))
}

fn clone_array(
    ctx: &mut vybe_runtime::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    let target_arc = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
    let target_val = Value::Object(target_arc.clone());
    seen.insert(id, target_val.clone());

    let (elems, holes): (Vec<Value>, Vec<usize>) = {
        let s = src.lock().unwrap();
        if let ObjectKind::Array(ref v) = s.kind {
            let holes = if v.is_empty() {
                Vec::new()
            } else {
                crate::array::hole_indices(&s)
                    .into_iter()
                    .filter(|index| *index < v.len())
                    .collect()
            };
            (v.clone(), holes)
        } else {
            (Vec::new(), Vec::new())
        }
    };
    {
        let mut t = target_arc.lock().unwrap();
        let cloned_elems: Result<Vec<Value>, Value> = elems
            .iter()
            .map(|e| deep_clone(ctx, e, seen))
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
        crate::array::store_hole_indices(&mut t, &hole_set);
    }
    Ok(target_val)
}

fn clone_map(
    ctx: &mut vybe_runtime::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    let mut target_obj = Object::new();
    target_obj.kind = ObjectKind::Map(indexmap::IndexMap::new());
    // ⛔ NO own `size`. §24.1.3.10 makes `size` an accessor on
    // %Map.prototype%, and `new_map_value` builds real Maps that way. An own
    // `size` data property here SHADOWS that accessor, so the clone reported
    // the size it had at clone time for ever: `clone.set("c", 3)` inserted
    // into `ObjectKind::Map` correctly and `clone.size` still answered 2.
    // Invisible while the clone had no prototype to shadow.
    // HTML structured clone creates the clone with the target realm's
    // %Map.prototype%. Without the link the clone answered `false` to
    // `instanceof Map` while still having working `get`/`set` — a working
    // object that is an instance of nothing.
    target_obj
        .properties
        .insert("__proto__".into(), crate::map::shared_map_prototype());
    target_obj
        .properties
        .insert("__type".into(), crate::keys::string_value("Map"));
    let target_arc = vybe_runtime::heap::alloc(target_obj);
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
        {
            if let ObjectKind::Map(ref mut m) = t.kind {
                if !entries.is_empty() {
                    m.reserve(entries.len());
                }
                for (k, v) in entries {
                    // Per spec: Map keys are cloned too (unlike JSON).
                    m.insert(
                        deep_clone(ctx, &k, seen)?,
                        deep_clone(ctx, &v, seen)?,
                    );
                }
            }
        }
    }
    Ok(target_val)
}

fn clone_set(
    ctx: &mut vybe_runtime::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    let mut target_obj = Object::new();
    target_obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    // ⛔ NO own `size` — see the Map clone above (§24.2.3.9 accessor).
    // See the Map clone above — %Set.prototype%.
    target_obj
        .properties
        .insert("__proto__".into(), crate::set::shared_set_prototype());
    target_obj
        .properties
        .insert("__type".into(), crate::keys::string_value("Set"));
    let target_arc = vybe_runtime::heap::alloc(target_obj);
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
        {
            if let ObjectKind::Set(ref mut set) = t.kind {
                if !elements.is_empty() {
                    set.reserve(elements.len());
                }
                for v in elements {
                    set.insert(deep_clone(ctx, &v, seen)?);
                }
            }
        }
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
            // Shared storage is reused, and detached storage is rejected.
            // Neither path needs a byte snapshot that will only be discarded.
            let bytes = if state.shared || state.detached {
                Vec::new()
            } else {
                state.bytes.lock().unwrap().clone()
            };
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
        return Err(crate::error::new_error_flat(
            "DataCloneError",
            "ArrayBuffer is detached",
        ));
    }
    if shared {
        let out = Value::Object(src.clone());
        seen.insert(id, out.clone());
        return Ok(out);
    }
    let byte_len = bytes_copy.len();
    let state = ArrayBufferState {
        bytes: Arc::new(Mutex::new(bytes_copy)),
        max_byte_length,
        resizable,
        detached: false,
        shared,
    };
    let mut obj = Object::new();
    obj.kind = ObjectKind::ArrayBuffer(state);
    obj.properties
        .insert("byteLength".into(), Value::I32(byte_len as i32));
    obj.properties
        .insert("maxByteLength".into(), Value::I32(max_byte_length as i32));
    obj.properties.insert(
        "__proto__".into(),
        crate::arraybuffer::shared_arraybuffer_prototype(),
    );
    obj.properties
        .insert("__type".into(), crate::keys::string_value("ArrayBuffer"));
    let out = Value::Object(vybe_runtime::heap::alloc(obj));
    seen.insert(id, out.clone());
    Ok(out)
}

fn clone_typedarray(
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
) -> Result<Value, Value> {
    let (elem, byte_offset, length, buffer_obj, buffer_detached) = {
        let s = src.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = s.kind {
            let detached = ta.buffer_obj.lock().unwrap().kind.as_arraybuffer_detached();
            (
                ta.elem,
                ta.byte_offset,
                ta.length,
                ta.buffer_obj.clone(),
                detached,
            )
        } else {
            return Ok(Value::Null);
        }
    };
    if buffer_detached {
        return Err(crate::error::new_error_flat(
            "DataCloneError",
            "ArrayBuffer is detached",
        ));
    }

    // Clone the backing buffer once per source identity, including for partial
    // views. Sibling views must remain views of that same cloned buffer, with
    // their original offsets and lengths, rather than independent byte copies.
    let buffer_id = Arc::as_ptr(&buffer_obj) as usize;
    let cloned_buffer = if let Some(Value::Object(existing)) = seen.get(&buffer_id) {
        existing.clone()
    } else {
        match clone_arraybuffer(&buffer_obj, buffer_id, seen)? {
            Value::Object(obj) => obj,
            _ => return Ok(Value::Null),
        }
    };
    let out = crate::typedarray::new_view_over_buffer(elem, cloned_buffer, byte_offset, length);
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
        .get(crate::arraybuffer::DV_TAG)
        .is_some_and(|v| !matches!(v, Value::Undefined | Value::Null | Value::Bool(false)))
}

fn is_uncloneable_weak_collection(s: &Object) -> bool {
    matches!(
        s.properties.get("__type"),
        Some(Value::String(tag)) if matches!(tag.as_ref(), "WeakMap" | "WeakSet")
    ) || s.properties.contains_key(crate::weakmap::WEAKMAP_TAG)
        || s.properties.contains_key(crate::weakmap::WEAKSET_TAG)
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
    {
        let b = buffer.lock().unwrap();
        if let ObjectKind::ArrayBuffer(ref state) = b.kind {
            if state.detached {
                return Err(crate::error::new_error_flat(
                    "DataCloneError",
                    "ArrayBuffer is detached",
                ));
            }
        }
    }
    let buffer_id = Arc::as_ptr(&buffer) as usize;
    let buffer_value = match seen.get(&buffer_id) {
        Some(existing) => existing.clone(),
        None => clone_arraybuffer(&buffer, buffer_id, seen)?,
    };
    let out = crate::arraybuffer::new_dataview(buffer_value, byte_offset as i32, byte_length as i32);
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
    ctx: &mut vybe_runtime::HostContext,
    src: &Arc<Mutex<Object>>,
    id: usize,
    seen: &mut HashMap<usize, Value>,
    source_kind: &str,
) -> Result<Value, Value> {
    let target_arc = vybe_runtime::heap::alloc(Object::new());
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
        // ⛔ And LINK it. `stamp_error_clone` writes `__type` / `name` /
        // `__types` — none of which makes the clone an instance of anything.
        // `structuredClone(new Error("x")) instanceof Error` was `false`.
        // Same per-VM `__ctor_<Kind>` anchor `error.rs::link_error_prototype`
        // uses; a realm with no wired error constructors keeps the stamps.
        if let Value::Object(ctor) = ctx.get_global(&format!("__ctor_{clone_kind}")) {
            let proto = ctor.lock().unwrap().properties.get("prototype").cloned();
            if let Some(proto @ Value::Object(_)) = proto {
                t.properties.insert("__proto__".into(), proto);
            }
        }
    }
    for (k, v) in entries {
        let cloned = deep_clone(ctx, &v, seen)?;
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
    ) || !crate::object::is_nonenum(o, key)
}

fn stamp_error_clone(obj: &mut Object, kind: &str, display_name: &str, message: &str) {
    obj.properties
        .insert("__type".into(), crate::keys::string_value(kind));
    obj.properties
        .insert("__exception_type".into(), crate::keys::string_value(kind));
    obj.properties
        .insert("name".into(), crate::keys::string_value(display_name));
    obj.properties
        .insert("message".into(), crate::keys::string_value(message));
    obj.properties.insert(
        "stack".into(),
        Value::String(crate::keys::concat3_arc(display_name, ": ", &message)),
    );
    let chain: Vec<Value> = error_ancestors(kind)
        .iter()
        .map(|n| crate::keys::string_value(n))
        .collect();
    obj.properties.insert(
        "__types".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(chain))),
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


#[cfg(test)]
mod graph_allocation_speedup_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn clone_value(value: &Value) -> Result<Value, Value> {
        deep_clone(&mut vybe_runtime::HostContext::empty(), value, &mut HashMap::new())
    }

    fn object(value: &Value) -> Arc<Mutex<Object>> {
        let Value::Object(value) = value else { panic!("expected object") };
        value.clone()
    }

    fn property(value: &Value, name: &str) -> Value {
        object(value).lock().unwrap().properties.get(name).unwrap().clone()
    }

    fn same_reference(a: &Value, b: &Value) -> bool {
        matches!((a, b), (Value::Object(a), Value::Object(b)) if Arc::ptr_eq(a, b))
    }

    fn buffer(size: usize, shared: bool, detached: bool) -> Arc<Mutex<Object>> {
        let mut value = Object::new();
        value.kind = ObjectKind::ArrayBuffer(ArrayBufferState {
            bytes: Arc::new(Mutex::new(vec![7; size])),
            max_byte_length: size,
            resizable: false,
            shared,
            detached,
        });
        vybe_runtime::heap::alloc(value)
    }

    #[test]
    fn one_seen_map_preserves_cycles_aliases_and_container_order() {
        let root = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let leaf = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        object(&leaf).lock().unwrap().properties.insert("n".into(), Value::I32(7));
        let array = Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            vec![root.clone(), leaf.clone(), leaf.clone()])));
        let mut map = Object::new();
        map.kind = ObjectKind::Map(indexmap::IndexMap::from([
            (root.clone(), leaf.clone()), (leaf.clone(), array.clone()),
        ]));
        let map = Value::Object(vybe_runtime::heap::alloc(map));
        let mut set = Object::new();
        set.kind = ObjectKind::Set(indexmap::IndexSet::from([root.clone(), leaf.clone()]));
        let set = Value::Object(vybe_runtime::heap::alloc(set));
        {
            let root_object = object(&root);
            let mut properties = root_object.lock().unwrap();
            for (key, value) in [("self", root.clone()), ("leaf", leaf.clone()),
                ("array", array), ("map", map), ("set", set)] {
                properties.properties.insert(key.into(), value);
            }
        }
        let cloned = clone_value(&root).unwrap();
        assert!(!same_reference(&cloned, &root));
        assert!(same_reference(&property(&cloned, "self"), &cloned));
        let cloned_leaf = property(&cloned, "leaf");
        assert!(!same_reference(&cloned_leaf, &leaf));
        assert_eq!(property(&cloned_leaf, "n").as_i32(), 7);
        let cloned_array = property(&cloned, "array");
        {
            let array_object = object(&cloned_array);
            let array = array_object.lock().unwrap();
            let ObjectKind::Array(elements) = &array.kind else { panic!("expected array") };
            assert_eq!(elements.len(), 3);
            assert!(same_reference(&elements[0], &cloned));
            assert!(same_reference(&elements[1], &cloned_leaf));
            assert!(same_reference(&elements[2], &cloned_leaf));
        }
        {
            let map_object = object(&property(&cloned, "map"));
            let map = map_object.lock().unwrap();
            let ObjectKind::Map(entries) = &map.kind else { panic!("expected map") };
            let ordered: Vec<_> = entries.iter().collect();
            assert_eq!(ordered.len(), 2);
            assert!(same_reference(ordered[0].0, &cloned));
            assert!(same_reference(ordered[0].1, &cloned_leaf));
            assert!(same_reference(ordered[1].0, &cloned_leaf));
            assert!(same_reference(ordered[1].1, &cloned_array));
        }
        {
            let set_object = object(&property(&cloned, "set"));
            let set = set_object.lock().unwrap();
            let ObjectKind::Set(entries) = &set.kind else { panic!("expected set") };
            let ordered: Vec<_> = entries.iter().collect();
            assert_eq!(ordered.len(), 2);
            assert!(same_reference(ordered[0], &cloned));
            assert!(same_reference(ordered[1], &cloned_leaf));
        }
    }

    #[test]
    fn ordinary_buffers_still_copy_bytes_and_shared_buffers_keep_existing_storage_behavior() {
        for shared in [false, true] {
            let source = buffer(32, shared, false);
            let cloned = clone_value(&Value::Object(source.clone())).unwrap();
            let cloned_object = object(&cloned);
            assert_eq!(Arc::ptr_eq(&source, &cloned_object), shared);
            let source_bytes = {
                let source = source.lock().unwrap();
                let ObjectKind::ArrayBuffer(state) = &source.kind else { panic!("expected buffer") };
                state.bytes.clone()
            };
            let cloned_bytes = {
                let clone = cloned_object.lock().unwrap();
                let ObjectKind::ArrayBuffer(state) = &clone.kind else { panic!("expected buffer") };
                assert_eq!(state.max_byte_length, 32);
                assert_eq!(state.shared, shared);
                assert!(!state.detached);
                state.bytes.clone()
            };
            assert_eq!(Arc::ptr_eq(&source_bytes, &cloned_bytes), shared);
            cloned_bytes.lock().unwrap()[0] = 99;
            assert_eq!(source_bytes.lock().unwrap()[0], if shared { 99 } else { 7 });
        }
        for shared in [false, true] {
            let detached = Value::Object(buffer(32, shared, true));
            assert!(clone_value(&detached).is_err());
        }
    }

    #[test]
    fn unsupported_nested_values_still_abort_cloning() {
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(0);
        let mut weak = Object::new();
        weak.properties.insert("__type".into(), crate::keys::string_value("WeakMap"));
        for unsupported in [
            Value::Symbol(crate::keys::string_arc("clone-symbol")),
            Value::Object(vybe_runtime::heap::alloc(function)),
            Value::Object(vybe_runtime::heap::alloc(weak)),
        ] {
            let nested = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![unsupported])));
            assert!(clone_value(&nested).is_err());
        }
    }

    fn host_ref(index: usize) -> Value {
        let mut value = Object::new();
        value.kind = ObjectKind::HostFunction(index);
        Value::Object(vybe_runtime::heap::alloc(value))
    }

    #[test]
    #[ignore = "paired registered ArrayBuffer helper benchmark, not whole structuredClone timing"]
    fn registered_buffer_copy_benchmark() {
        let mut vm = VM::new();
        vm.register_free_fn("test:buffer-clone", "previous", Box::new(|_, args| {
            let source = object(&args[0]);
            previous_clone_arraybuffer(&source, Arc::as_ptr(&source) as usize,
                &mut HashMap::new()).unwrap()
        }));
        vm.register_free_fn("test:buffer-clone", "optimized", Box::new(|_, args| {
            let source = object(&args[0]);
            clone_arraybuffer(&source, Arc::as_ptr(&source) as usize,
                &mut HashMap::new()).unwrap()
        }));
        let previous = host_ref(vm.resolve_host_function_index("test:buffer-clone", "previous").unwrap());
        let optimized = host_ref(vm.resolve_host_function_index("test:buffer-clone", "optimized").unwrap());
        for shared in [false, true] {
            for size in [1024, 65536, 4194304] {
                let args = [Value::Object(buffer(size, shared, false))];
                let calls = (8 * 1024 * 1024 / size).clamp(8, 2048);
                let mut old = Duration::ZERO;
                let mut new = Duration::ZERO;
                for batch in 0..8 {
                    for old_path in if batch % 2 == 0 { [true, false] } else { [false, true] } {
                        let callee = if old_path { &previous } else { &optimized };
                        let start = Instant::now();
                        for _ in 0..calls {
                            black_box(vm.invoke(black_box(callee), black_box(&args)).unwrap());
                        }
                        if old_path { old += start.elapsed(); } else { new += start.elapsed(); }
                    }
                }
                eprintln!("registered ArrayBuffer clone helper, shared={shared}, bytes={size}, calls={}: previous={old:?}, optimized={new:?}, ratio={:.3}x",
                    calls * 8, old.as_secs_f64() / new.as_secs_f64());
            }
        }
    }
    fn previous_clone_arraybuffer(
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
            return Err(crate::error::new_error_flat(
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
        obj.properties.insert(
            "__proto__".into(),
            crate::arraybuffer::shared_arraybuffer_prototype(),
        );
        obj.properties
            .insert("__type".into(), crate::keys::string_value("ArrayBuffer"));
        let out = Value::Object(vybe_runtime::heap::alloc(obj));
        seen.insert(id, out.clone());
        Ok(out)
    }
    
}

fn string_prop(o: &Object, key: &str) -> Option<String> {
    match o.properties.get(key) {
        Some(Value::String(s)) => Some(s.to_string()),
        Some(v) => Some(crate::keys::value_display_string(v)),
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
    _ctx: &vybe_runtime::HostContext,
    options: Option<&Value>,
) -> Option<Value> {
    match options {
        None | Some(Value::Undefined) => None,
        Some(Value::Object(_)) => None,
        Some(_) => Some(crate::error::new_error_flat(
            "TypeError",
            "structuredClone options must be an object",
        )),
    }
}

fn collect_transfer_list(
    ctx: &mut vybe_runtime::HostContext,
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
                    return Err(crate::error::new_error_flat(
                        "TypeError",
                        "transfer option is not iterable",
                    ));
                }
            }
        }
        Value::Null => {
            return Err(crate::error::new_error_flat(
                "TypeError",
                "transfer list contains null",
            ));
        }
        _ => {
            return Err(crate::error::new_error_flat(
                "TypeError",
                "transfer option is not iterable",
            ));
        }
    }
    let mut ids = HashSet::new();
    for item in &out {
        let id = Arc::as_ptr(item) as usize;
        if !ids.insert(id) {
            return Err(crate::error::new_error_flat(
                "DataCloneError",
                "duplicate transfer item",
            ));
        }
    }
    Ok(out)
}

fn push_transferable(value: &Value, out: &mut Vec<Arc<Mutex<Object>>>) -> Result<(), Value> {
    let Value::Object(obj) = value else {
        return Err(crate::error::new_error_flat(
            "DataCloneError",
            "transfer item is not transferable",
        ));
    };
    {
        let o = obj.lock().unwrap();
        let ObjectKind::ArrayBuffer(state) = &o.kind else {
            return Err(crate::error::new_error_flat(
                "DataCloneError",
                "transfer item is not transferable",
            ));
        };
        if state.shared {
            return Err(crate::error::new_error_flat(
                "DataCloneError",
                "SharedArrayBuffer is not transferable",
            ));
        }
        if state.detached {
            return Err(crate::error::new_error_flat(
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
    // Without transfers there are no detached views to update. Avoid walking
    // the entire source graph (and cloning its child handles) in this case.
    if transferred.is_empty() {
        return;
    }
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
        } else if o.properties.contains_key(crate::arraybuffer::DV_TAG) {
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

#[cfg(test)]
mod transfer_walk_speedup_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn graph(depth: usize) -> Value {
        let items = if depth == 0 {
            vec![Value::I32(7)]
        } else {
            vec![graph(depth - 1), graph(depth - 1)]
        };
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
    }

    #[test]
    fn empty_transfer_list_does_not_visit_source_graph() {
        let value = graph(4);
        let mut seen = HashSet::new();
        mark_transferred_views(&value, &[], &mut seen);
        assert!(seen.is_empty());
        let Value::Object(root) = value else { panic!("expected array") };
        let root = root.lock().unwrap();
        assert!(matches!(&root.kind, ObjectKind::Array(items) if items.len() == 2));
    }

    #[test]
    fn nonempty_transfer_list_still_visits_graph() {
        let value = graph(4);
        let item = vybe_runtime::heap::alloc(Object::new());
        let mut seen = HashSet::new();
        mark_transferred_views(&value, &[item], &mut seen);
        assert_eq!(seen.len(), 31);
    }

    #[test]
    #[ignore = "native helper microbenchmark, not whole structuredClone timing"]
    fn native_empty_transfer_walk_microbenchmark() {
        let value = graph(9);
        let ids = HashSet::new();
        let mut old = Duration::ZERO;
        let mut new = Duration::ZERO;
        for batch in 0..10 {
            for old_path in if batch % 2 == 0 { [true, false] } else { [false, true] } {
                let start = Instant::now();
                for _ in 0..100 {
                    let mut seen = HashSet::new();
                    if old_path {
                        mark_transferred_views_inner(black_box(&value), &ids, &mut seen);
                        assert_eq!(seen.len(), 1023);
                    } else {
                        mark_transferred_views(black_box(&value), black_box(&[]), &mut seen);
                        assert!(seen.is_empty());
                    }
                    black_box(seen);
                }
                if old_path { old += start.elapsed(); } else { new += start.elapsed(); }
            }
        }
        eprintln!("empty transfer post-pass, 1000 calls/1023 nodes: original={old:?}, skipped={new:?}");
    }
}

#[cfg(test)]
mod dataview_clone_speedup_tests {
    use super::*;

    fn source_buffer(detached: bool) -> Arc<Mutex<Object>> {
        let mut obj = Object::new();
        obj.kind = ObjectKind::ArrayBuffer(ArrayBufferState {
            bytes: Arc::new(Mutex::new((0..32).collect())),
            max_byte_length: 32,
            resizable: false,
            detached,
            shared: false,
        });
        vybe_runtime::heap::alloc(obj)
    }

    fn view_parts(view: &Value) -> (Arc<Mutex<Object>>, i32, i32) {
        let Value::Object(view) = view else { panic!("expected DataView") };
        let view = view.lock().unwrap();
        let Some(Value::Object(buffer)) = view.properties.get("buffer")
            else { panic!("expected backing buffer") };
        (buffer.clone(), view.properties.get("byteOffset").unwrap().as_i32(),
            view.properties.get("byteLength").unwrap().as_i32())
    }

    fn clone_view(view: &Value, seen: &mut HashMap<usize, Value>) -> Value {
        let Value::Object(view) = view else { panic!("expected DataView") };
        clone_dataview(view, Arc::as_ptr(view) as usize, seen).unwrap()
    }

    #[test]
    fn sibling_views_preserve_offsets_lengths_full_backing_and_overlap() {
        let source = source_buffer(false);
        let a = crate::arraybuffer::new_dataview(Value::Object(source.clone()), 4, 8);
        let b = crate::arraybuffer::new_dataview(Value::Object(source.clone()), 8, 8);
        let mut seen = HashMap::new();
        let ca = clone_view(&a, &mut seen);
        let cb = clone_view(&b, &mut seen);
        let (ab, ao, al) = view_parts(&ca);
        let (bb, bo, bl) = view_parts(&cb);
        assert_eq!((ao, al, bo, bl), (4, 8, 8, 8));
        assert!(Arc::ptr_eq(&ab, &bb));
        assert!(!Arc::ptr_eq(&ab, &source));
        let Some(Value::Object(mapped)) = seen.get(&(Arc::as_ptr(&source) as usize))
            else { panic!("expected mapped backing buffer") };
        assert!(Arc::ptr_eq(mapped, &ab));
        {
            let buffer = ab.lock().unwrap();
            let ObjectKind::ArrayBuffer(state) = &buffer.kind else { panic!("expected ArrayBuffer") };
            let mut bytes = state.bytes.lock().unwrap();
            assert_eq!(bytes.len(), 32);
            assert_eq!(bytes[4], 4);
            bytes[8] = 199;
        }
        {
            let buffer = bb.lock().unwrap();
            let ObjectKind::ArrayBuffer(state) = &buffer.kind else { panic!("expected ArrayBuffer") };
            assert_eq!(state.bytes.lock().unwrap()[bo as usize], 199);
        }
        let buffer = source.lock().unwrap();
        let ObjectKind::ArrayBuffer(state) = &buffer.kind else { panic!("expected ArrayBuffer") };
        assert_eq!(state.bytes.lock().unwrap()[8], 8);
    }

    #[test]
    fn buffer_first_order_and_empty_view_keep_the_existing_clone() {
        let source = source_buffer(false);
        let mut seen = HashMap::new();
        let cloned = clone_arraybuffer(&source, Arc::as_ptr(&source) as usize, &mut seen).unwrap();
        let view = crate::arraybuffer::new_dataview(Value::Object(source), 12, 0);
        let result = clone_view(&view, &mut seen);
        let (buffer, offset, length) = view_parts(&result);
        let Value::Object(cloned) = cloned else { panic!("expected ArrayBuffer") };
        assert!(Arc::ptr_eq(&buffer, &cloned));
        assert_eq!((offset, length), (12, 0));
    }

    #[test]
    fn detached_backing_is_rejected_without_inserting_a_clone() {
        let source = source_buffer(false);
        let view = crate::arraybuffer::new_dataview(Value::Object(source.clone()), 4, 8);
        if let ObjectKind::ArrayBuffer(state) = &mut source.lock().unwrap().kind {
            state.detached = true;
            state.bytes.lock().unwrap().clear();
        }
        let mut seen = HashMap::new();
        let Value::Object(view) = view else { panic!("expected DataView") };
        assert!(clone_dataview(&view, Arc::as_ptr(&view) as usize, &mut seen).is_err());
        assert!(seen.is_empty());
    }
}

#[cfg(test)]
mod collection_capacity_speedup_tests {
    use super::*;

    fn object(value: &Value) -> Arc<Mutex<Object>> {
        let Value::Object(object) = value else { panic!("expected object") };
        object.clone()
    }

    fn clone_value(value: &Value) -> Value {
        deep_clone(&mut vybe_runtime::HostContext::empty(), value, &mut HashMap::new()).unwrap()
    }

    #[test]
    fn reserved_map_keeps_order_aliases_and_self_cycle() {
        let leaf = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let mut source = Object::new();
        source.kind = ObjectKind::Map((0..32).map(|index| (Value::I32(index), leaf.clone())).collect());
        let source = Value::Object(vybe_runtime::heap::alloc(source));
        {
            let source_object = object(&source);
            let mut source_object = source_object.lock().unwrap();
            let ObjectKind::Map(entries) = &mut source_object.kind else { unreachable!() };
            entries.insert(Value::I32(-1), source.clone());
        }
        let cloned = clone_value(&source);
        let cloned_object = object(&cloned);
        assert!(!Arc::ptr_eq(&object(&source), &cloned_object));
        let cloned_guard = cloned_object.lock().unwrap();
        let ObjectKind::Map(entries) = &cloned_guard.kind else { panic!("expected map") };
        assert_eq!(entries.len(), 33);
        let cloned_leaf = object(entries.get_index(0).unwrap().1);
        assert!(!Arc::ptr_eq(&cloned_leaf, &object(&leaf)));
        for index in 0..32 {
            let (key, value) = entries.get_index(index).unwrap();
            assert!(matches!(key, Value::I32(value) if *value == index as i32));
            assert!(Arc::ptr_eq(&object(value), &cloned_leaf));
        }
        let (key, value) = entries.get_index(32).unwrap();
        assert!(matches!(key, Value::I32(-1)));
        assert!(Arc::ptr_eq(&object(value), &cloned_object));
    }

    #[test]
    fn reserved_set_keeps_order_self_cycle_and_empty_collections() {
        let mut source = Object::new();
        source.kind = ObjectKind::Set((0..128).map(Value::I32).collect());
        let source = Value::Object(vybe_runtime::heap::alloc(source));
        {
            let source_object = object(&source);
            let mut source_object = source_object.lock().unwrap();
            let ObjectKind::Set(entries) = &mut source_object.kind else { unreachable!() };
            entries.insert(source.clone());
        }
        let cloned = clone_value(&source);
        let cloned_object = object(&cloned);
        {
            let cloned_guard = cloned_object.lock().unwrap();
            let ObjectKind::Set(entries) = &cloned_guard.kind else { panic!("expected set") };
            assert_eq!(entries.len(), 129);
            for index in 0..128 {
                assert!(matches!(entries.get_index(index), Some(Value::I32(value)) if *value == index as i32));
            }
            assert!(Arc::ptr_eq(&object(entries.get_index(128).unwrap()), &cloned_object));
        }
        for kind in [ObjectKind::Map(indexmap::IndexMap::new()), ObjectKind::Set(indexmap::IndexSet::new())] {
            let mut empty = Object::new();
            empty.kind = kind;
            let empty = Value::Object(vybe_runtime::heap::alloc(empty));
            let cloned = clone_value(&empty);
            assert!(!Arc::ptr_eq(&object(&empty), &object(&cloned)));
            let cloned = object(&cloned);
            let cloned = cloned.lock().unwrap();
            match &cloned.kind {
                ObjectKind::Map(entries) => assert!(entries.is_empty()),
                ObjectKind::Set(entries) => assert!(entries.is_empty()),
                _ => panic!("collection kind changed"),
            }
        }
    }

    #[test]
    #[ignore = "paired target collection insertion benchmark, excludes recursive clone and snapshots"]
    fn paired_capacity_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        for length in [0, 8, 64, 512, 4096] {
            let entries: Vec<_> = (0..length).map(|index| {
                (crate::keys::owned_string_value(format!("key-{index}")), Value::I32(index as i32))
            }).collect();
            let iterations = if length > 512 { 20 } else if length > 64 { 200 } else { 1000 };
            for map in [true, false] {
                let mut previous = Duration::ZERO;
                let mut reserved = Duration::ZERO;
                for round in 0..8 {
                    for reserve in if round % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            if map {
                                let mut target = indexmap::IndexMap::new();
                                if reserve && !entries.is_empty() { target.reserve(entries.len()); }
                                for (key, value) in black_box(&entries) { target.insert(key.clone(), value.clone()); }
                                black_box(target);
                            } else {
                                let mut target = indexmap::IndexSet::new();
                                if reserve && !entries.is_empty() { target.reserve(entries.len()); }
                                for (key, _) in black_box(&entries) { target.insert(key.clone()); }
                                black_box(target);
                            }
                        }
                        if reserve { reserved += start.elapsed(); } else { previous += start.elapsed(); }
                    }
                }
                eprintln!("clone target {} len={length}: previous={previous:?}, reserved={reserved:?}", if map { "map" } else { "set" });
            }
        }
    }
}

#[cfg(test)]
mod ordinary_snapshot_speedup_tests {
    use super::*;

    fn object(value: &Value) -> Arc<Mutex<Object>> {
        let Value::Object(object) = value else { panic!("expected object") };
        object.clone()
    }

    fn named(vm: &VM, name: &str) -> Value {
        let mut value = Object::new();
        value.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("test:clone-snapshot", name).unwrap());
        Value::Object(vybe_runtime::heap::alloc(value))
    }

    #[test]
    fn getters_keep_order_source_reentry_aliases_and_cycles() {
        let mut vm = VM::new();
        register(&mut vm);
        let source = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let leaf = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let trace = Arc::new(Mutex::new(Vec::new()));
        for index in 1..=3 {
            let source_object = object(&source);
            let source_value = source.clone();
            let leaf = leaf.clone();
            let trace = trace.clone();
            vm.register_free_fn("test:clone-snapshot", &format!("getter-{index}"), Box::new(move |_, args| {
                assert_eq!(args.len(), 1);
                assert!(matches!(&args[0], Value::Object(actual) if Arc::ptr_eq(actual, &source_object)));
                let mut source = source_object.try_lock().expect("source guard held across getter");
                source.properties.insert("data".into(), Value::I32(99));
                drop(source);
                trace.lock().unwrap().push(index);
                if index == 3 { source_value.clone() } else { leaf.clone() }
            }));
        }
        {
            let source_object = object(&source);
            let mut source = source_object.lock().unwrap();
            source.properties.insert("data".into(), Value::I32(7));
            source.properties.insert("leaf".into(), leaf.clone());
            source.properties.insert("__get_first".into(), named(&vm, "getter-1"));
            source.properties.insert("late".into(), Value::I32(13));
            source.properties.insert("__get_second".into(), named(&vm, "getter-2"));
            source.properties.insert("__get_self".into(), named(&vm, "getter-3"));
        }
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:structured-clone", "clone").unwrap());
        let function = Value::Object(vybe_runtime::heap::alloc(function));
        let cloned = vm.invoke(&function, &[source.clone()]).unwrap();
        assert_eq!(*trace.lock().unwrap(), vec![1, 2, 3]);
        let cloned_object = object(&cloned);
        let cloned_guard = cloned_object.lock().unwrap();
        assert!(matches!(cloned_guard.properties.get("data"), Some(Value::I32(7))));
        let cloned_leaf = object(cloned_guard.properties.get("leaf").unwrap());
        assert!(!Arc::ptr_eq(&cloned_leaf, &object(&leaf)));
        for name in ["first", "second"] {
            assert!(Arc::ptr_eq(&cloned_leaf, &object(cloned_guard.properties.get(name).unwrap())));
        }
        assert!(Arc::ptr_eq(&cloned_object, &object(cloned_guard.properties.get("self").unwrap())));
        let names: Vec<_> = cloned_guard.properties.keys().map(String::as_str).collect();
        assert_eq!(names, ["data", "leaf", "late", "first", "second", "self"]);
    }

    fn previous_snapshots(source: &Object) -> (Vec<(String, Value)>, Vec<(String, Value)>) {
        let entries = source.properties.iter().filter(|(key, _)| {
            !key.starts_with("__") && !key.starts_with("Symbol(") && !crate::object::is_nonenum(source, key)
        }).map(|(key, value)| (key.clone(), value.clone())).collect();
        let accessors = source.properties.iter().filter_map(|(key, getter)| {
            let property = key.strip_prefix("__get_")?;
            if property.starts_with("__") || property.starts_with("Symbol(") || crate::object::is_nonenum(source, property) {
                return None;
            }
            Some((property.to_string(), getter.clone()))
        }).collect();
        (entries, accessors)
    }

    #[test]
    fn fused_snapshot_keeps_existing_filtering_and_list_order() {
        let mut source = Object::new();
        for (name, value) in [("a", 1), ("__get_b", 2), ("c", 3), ("__hidden", 4),
            ("Symbol(key)", 5), ("__get_Symbol(key)", 6), ("__get___hidden", 7)] {
            source.properties.insert(name.into(), Value::I32(value));
        }
        let old = previous_snapshots(&source);
        let new = ordinary_property_snapshots(&source);
        for (old, new) in [(old.0, new.0), (old.1, new.1)] {
            assert_eq!(old.len(), new.len());
            for ((old_key, old_value), (new_key, new_value)) in old.into_iter().zip(new) {
                assert_eq!(old_key, new_key);
                assert!(matches!((old_value, new_value), (Value::I32(a), Value::I32(b)) if a == b));
            }
        }
    }

    #[test]
    #[ignore = "paired property snapshot benchmark, excludes recursive clone and getters"]
    fn paired_snapshot_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        for length in [0, 8, 64, 512] {
            let mut source = Object::new();
            for index in 0..length {
                let name = if index % 4 == 0 { format!("__get_key-{index}") } else { format!("key-{index}") };
                source.properties.insert(name, Value::I32(index));
            }
            let mut previous = Duration::ZERO;
            let mut fused = Duration::ZERO;
            for round in 0..8 {
                for optimized in if round % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..1000 {
                        black_box(if optimized { ordinary_property_snapshots(black_box(&source)) }
                            else { previous_snapshots(black_box(&source)) });
                    }
                    if optimized { fused += start.elapsed(); } else { previous += start.elapsed(); }
                }
            }
            eprintln!("ordinary property snapshot len={length}: previous={previous:?}, fused={fused:?}");
        }
    }
}

#[cfg(test)]
mod array_hole_snapshot_speedup_tests {
    use super::*;

    fn previous_holes(source: &Object, length: usize) -> Vec<usize> {
        (0..length).filter(|index| crate::array::is_array_hole(source, *index)).collect()
    }

    fn current_holes(source: &Object, length: usize) -> Vec<usize> {
        if length == 0 {
            return Vec::new();
        }
        crate::array::hole_indices(source).into_iter().filter(|index| *index < length).collect()
    }

    #[test]
    fn hole_snapshot_matches_index_checks_for_dense_sparse_and_bounds() {
        for length in [0, 1, 8, 64] {
            let mut source = Object::new_array(vec![Value::Undefined; length]);
            assert_eq!(current_holes(&source, length), previous_holes(&source, length));
            let holes: BTreeSet<_> = [0, 3, 7, 63, 1000].into_iter().collect();
            crate::array::store_hole_indices(&mut source, &holes);
            assert_eq!(current_holes(&source, length), previous_holes(&source, length));
            crate::array::store_hole_indices(&mut source, &BTreeSet::new());
            assert!(current_holes(&source, length).is_empty());
        }
    }

    #[test]
    fn snapshot_preserves_existing_integer_and_invalid_metadata_handling() {
        let mut source = Object::new_array(vec![Value::Undefined; 8]);
        let metadata = Object::new_array(vec![
            Value::I32(-1), Value::I32(0), Value::I32(0), Value::I64(3), Value::I64(-2),
            Value::I64(1_i64 << 32), Value::F64(2.0), Value::Null, Value::Undefined,
        ]);
        source.properties.insert("__holes".into(), Value::Object(vybe_runtime::heap::alloc(metadata)));
        assert_eq!(current_holes(&source, 8), previous_holes(&source, 8));
        assert_eq!(current_holes(&source, 8), vec![0, 3]);
        for metadata in [Value::Bool(true), Value::Null, Value::Object(vybe_runtime::heap::alloc(Object::new()))] {
            source.properties.insert("__holes".into(), metadata);
            assert_eq!(current_holes(&source, 8), previous_holes(&source, 8));
        }
    }

    #[test]
    fn cloned_sparse_arrays_keep_holes_explicit_undefined_and_aliases() {
        let leaf = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let mut source = Object::new_array(vec![leaf.clone(), Value::Undefined, leaf.clone(), Value::Undefined]);
        crate::array::store_hole_indices(&mut source, &[3].into_iter().collect());
        let source = Value::Object(vybe_runtime::heap::alloc(source));
        let cloned = deep_clone(&mut vybe_runtime::HostContext::empty(), &source, &mut HashMap::new()).unwrap();
        let Value::Object(cloned) = cloned else { panic!("expected array") };
        let cloned = cloned.lock().unwrap();
        assert!(!crate::array::is_array_hole(&cloned, 1));
        assert!(crate::array::is_array_hole(&cloned, 3));
        let ObjectKind::Array(values) = &cloned.kind else { panic!("expected array backing") };
        assert_eq!(values.len(), 4);
        assert!(matches!(&values[1], Value::Undefined));
        assert!(matches!((&values[0], &values[2]), (Value::Object(a), Value::Object(b)) if Arc::ptr_eq(a, b)));
    }

    #[test]
    #[ignore = "paired hole snapshot benchmark, excludes element cloning and target allocation"]
    fn paired_hole_snapshot_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        for length in [0, 64, 512, 4096] {
            for sparse in [false, true] {
                let mut source = Object::new_array(vec![Value::Undefined; length]);
                if sparse {
                    let holes = (0..length).step_by(4).collect();
                    crate::array::store_hole_indices(&mut source, &holes);
                }
                let iterations = if length > 512 { 2 } else if length > 64 { 25 } else { 100 };
                let mut previous = Duration::ZERO;
                let mut snapshot = Duration::ZERO;
                for round in 0..8 {
                    for current in if round % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..iterations {
                            black_box(if current { current_holes(black_box(&source), length) }
                                else { previous_holes(black_box(&source), length) });
                        }
                        if current { snapshot += start.elapsed(); } else { previous += start.elapsed(); }
                    }
                }
                eprintln!("array holes len={length} sparse={sparse}: previous={previous:?}, snapshot={snapshot:?}");
            }
        }
    }
}
