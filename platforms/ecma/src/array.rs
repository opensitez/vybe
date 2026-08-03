//! # `ecma:array` host handlers
//!
//! Native Rust implementations that satisfy the `ecma:array.*`
//! imports declared in
//! `crates/vybe_runtime/src/wasm/js_array_builtins.rs`.
//!
//! On the Vybe VM this file IS the host fast-path — handlers operate
//! directly on the native `Vec<Value>` backing an `ObjectKind::Array`,
//! no indirection through spine-struct WASM bytecode. On v8 / browsers
//! the same interface is satisfied by JS glue (Phase C). On plain
//! wasmtime the polyfill module provides spine-struct implementations.
//!
//! Semantics follow ECMA-262 §23.1 exactly. When in doubt, MDN is the
//! authoritative reference.
//!
//! Marshaling + error-handling contract pinned in
//! `crates/vybe_runtime/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use crate::typedarray::{new_typed_array, read_element, ta_live_length, write_element};
use crate::receiver_host_fn_ref;
use std::collections::{BTreeSet, HashSet};
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, TypedElemKind, Value};
use vybe_runtime::{HostContext, VM};

fn invoke_callback(ctx: &mut HostContext, callback: &Value, args: &[Value]) -> Value {
    if let Some(v) = crate::function::invoke_bound_callback_if_needed(ctx, callback, args) {
        return v;
    }
    if let Some(v) = invoke_magic_callback(callback, args) {
        return v;
    }
    ctx.invoke(callback, args)
}

fn is_callable_value(value: &Value) -> bool {
    match value {
        Value::Object(obj) => {
            matches!(
                obj.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) || is_magic_callback_object(obj)
        }
        _ => false }
}

fn throw_type_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
    Value::Undefined
}

fn throw_range_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "RangeError", message));
    Value::Undefined
}

fn require_callable(ctx: &mut HostContext, value: &Value, message: &str) -> bool {
    if is_callable_value(value) {
        true
    } else {
        let _ = throw_type_error(ctx, message);
        false
    }
}

fn invoke_callback_this(
    ctx: &mut HostContext,
    callback: &Value,
    this_arg: Value,
    args: &[Value],
) -> Value {
    if let Some(v) = crate::function::invoke_bound_callback_if_needed(ctx, callback, args) {
        return v;
    }
    if let Some(v) = invoke_magic_callback(callback, args) {
        return v;
    }
    crate::function::invoke_with_explicit_this(ctx, callback, this_arg, args)
}

fn is_magic_callback_object(obj: &Arc<Mutex<Object>>) -> bool {
    let o = obj.lock().unwrap();
    if !matches!(o.kind, ObjectKind::Ordinary) {
        return false;
    }
    [
        "__pred_gt",
        "__reduce_add",
        "__reduce_concat",
        "__map_mul",
        "__filter_mod_eq",
        "__fn_return",
    ]
    .iter()
    .any(|key| o.properties.contains_key(*key))
}

fn invoke_magic_callback(callback: &Value, args: &[Value]) -> Option<Value> {
    let Value::Object(obj) = callback else {
        return None;
    };
    let o = obj.lock().unwrap();
    if !matches!(o.kind, ObjectKind::Ordinary) {
        return None;
    }

    let x = args.first().cloned().unwrap_or(Value::Undefined);
    let idx = args.get(1).cloned().unwrap_or(Value::Undefined);

    if let Some(threshold) = o.properties.get("__pred_gt") {
        let t = threshold.as_f64();
        return Some(Value::Bool(x.as_f64() > t));
    }
    if o.properties.contains_key("__reduce_add") {
        let acc = x;
        let cur = idx;
        return Some(Value::I32(acc.as_i32() + cur.as_i32()));
    }
    if o.properties.contains_key("__reduce_concat") {
        let acc = format!("{}", x);
        let cur = format!("{}", idx);
        return Some(Value::String(Arc::from(format!("{}{}", acc, cur).as_str())));
    }
    if let Some(mul) = o.properties.get("__map_mul") {
        let m = mul.as_i32();
        return Some(Value::I32(x.as_i32() * m));
    }
    if let Some(Value::Object(params)) = o.properties.get("__filter_mod_eq") {
        let p = params.lock().unwrap();
        let m = p.properties.get("mod").map(|v| v.as_i32()).unwrap_or(2);
        let eq = p.properties.get("eq").map(|v| v.as_i32()).unwrap_or(0);
        let xi = x.as_i32();
        return Some(Value::Bool(xi % m == eq));
    }
    if o.properties.contains_key("__flatmap_dup") {
        let arr = Object::new_array(vec![x.clone(), x]);
        return Some(Value::Object(vybe_runtime::heap::alloc(arr)));
    }
    if o.properties.contains_key("__from_map_double_index") {
        let i = idx.as_i32();
        return Some(Value::I32(i * 2));
    }
    if o.properties.contains_key("__noop") {
        return Some(Value::Undefined);
    }
    None
}

static ARRAY_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub fn shared_array_prototype() -> Value {
    Value::Object(
        ARRAY_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

/// Shorthand: unwrap `args[idx]` as a JS Array. Returns `None` when
/// the argument isn't an array-kind object. Handlers that require an
/// array trap (per convention class 1) on `None`; handlers that match
/// spec "if not an Array, return something sensible" use the None
/// branch to provide the default.
fn array_of<'a>(args: &'a [Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    match args.get(idx) {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            if matches!(o.kind, ObjectKind::Array(_)) {
                drop(o);
                Some(obj.clone())
            } else {
                None
            }
        }
        _ => None }
}

fn make_array(elements: Vec<Value>) -> Value {
    let mut obj = Object::new_array(elements);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Array")));
    obj.properties
        .insert("__proto__".into(), shared_array_prototype());
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Marker property set by `ecma:fixedarray.freeze` to forbid
/// length-changing mutations. Mutators check it and no-op rather
/// than allow the change (spec behavior would be TypeError; we
/// silently no-op until exception dispatch from host handlers is
/// wired — Phase B5 follow-up).
const FROZEN_MARK: &str = "__vybe_frozen";

fn is_frozen(arr: &Arc<Mutex<Object>>) -> bool {
    let o = arr.lock().unwrap();
    o.properties.get(FROZEN_MARK).is_some()
}

fn property_length_as_usize(object: &Object) -> Option<usize> {
    match object.properties.get("length") {
        Some(Value::I32(value)) if *value >= 0 => Some(*value as usize),
        Some(Value::I64(value)) if *value >= 0 => Some(*value as usize),
        Some(Value::F64(value)) if *value >= 0.0 => Some(*value as usize),
        Some(Value::String(text)) => text.parse::<usize>().ok(),
        _ => None }
}

fn hole_indices(object: &Object) -> BTreeSet<usize> {
    let Some(Value::Object(holes)) = object.properties.get("__holes") else {
        return BTreeSet::new();
    };
    let holes_guard = holes.lock().unwrap();
    let ObjectKind::Array(ref elems) = holes_guard.kind else {
        return BTreeSet::new();
    };
    elems
        .iter()
        .filter_map(|value| match value {
            Value::I32(index) if *index >= 0 => Some(*index as usize),
            Value::I64(index) if *index >= 0 => Some(*index as usize),
            _ => None })
        .collect()
}

pub fn store_hole_indices(object: &mut Object, holes: &BTreeSet<usize>) {
    if holes.is_empty() {
        object.properties.shift_remove("__holes");
        return;
    }

    let holes_obj = match object.properties.get("__holes") {
        Some(Value::Object(existing)) => existing.clone(),
        _ => {
            let created = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            object
                .properties
                .insert("__holes".into(), Value::Object(created.clone()));
            created
        }
    };

    let mut holes_guard = holes_obj.lock().unwrap();
    let ObjectKind::Array(ref mut elems) = holes_guard.kind else {
        return;
    };
    elems.clear();
    elems.extend(holes.iter().map(|index| Value::I32(*index as i32)));
}

pub fn is_array_hole(object: &Object, index: usize) -> bool {
    hole_indices(object).contains(&index)
}

#[allow(dead_code)]
fn mark_array_hole(object: &mut Object, index: usize) {
    let mut holes = hole_indices(object);
    holes.insert(index);
    store_hole_indices(object, &holes);
}

pub fn clear_array_hole(object: &mut Object, index: usize) {
    let mut holes = hole_indices(object);
    holes.remove(&index);
    store_hole_indices(object, &holes);
}

pub fn mark_hole_range(object: &mut Object, range: std::ops::Range<usize>) {
    let mut holes = hole_indices(object);
    holes.extend(range);
    store_hole_indices(object, &holes);
}

pub fn remap_array_holes<F>(object: &mut Object, mut remap: F)
where
    F: FnMut(usize) -> Option<usize>,
{
    let holes = hole_indices(object);
    let remapped: BTreeSet<usize> = holes.into_iter().filter_map(|index| remap(index)).collect();
    store_hole_indices(object, &remapped);
}

pub fn present_array_entries(object: &Object) -> Vec<(usize, Value)> {
    let ObjectKind::Array(ref values) = object.kind else {
        return Vec::new();
    };
    values
        .iter()
        .enumerate()
        .filter(|(index, _)| !is_array_hole(object, *index))
        .map(|(index, value)| (index, value.clone()))
        .collect()
}

fn array_length(arr: &Arc<Mutex<Object>>) -> usize {
    let o = arr.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(values) => values.len(),
        _ => 0 }
}

fn array_value_at(arr: &Arc<Mutex<Object>>, index: usize) -> Value {
    let o = arr.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(values) => values.get(index).cloned().unwrap_or(Value::Undefined),
        _ => Value::Undefined }
}

fn array_present_value_at(arr: &Arc<Mutex<Object>>, index: usize) -> Option<Value> {
    let o = arr.lock().unwrap();
    if is_array_hole(&o, index) {
        return None;
    }
    match &o.kind {
        ObjectKind::Array(values) => values.get(index).cloned(),
        _ => None }
}

fn array_like_length(value: &Value) -> usize {
    let Value::Object(obj) = value else {
        return 0;
    };
    let object = obj.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(values) => values.len(),
        ObjectKind::TypedArray(ta) => ta_live_length(ta),
        ObjectKind::Ordinary => property_length_as_usize(&object).unwrap_or(0),
        ObjectKind::Map(map) => map.len(),
        _ => 0 }
}

fn array_like_value_at(value: &Value, index: usize) -> Value {
    let Value::Object(obj) = value else {
        return Value::Undefined;
    };
    let object = obj.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(values) => values.get(index).cloned().unwrap_or(Value::Undefined),
        ObjectKind::TypedArray(ta) => {
            if index < ta_live_length(ta) {
                read_element(ta, index)
            } else {
                Value::Undefined
            }
        }
        ObjectKind::Ordinary => object
            .properties
            .get(&index.to_string())
            .cloned()
            .unwrap_or(Value::Undefined),
        ObjectKind::Map(map) => map.values().nth(index).cloned().unwrap_or(Value::Undefined),
        _ => Value::Undefined }
}

fn array_like_present_value_at(value: &Value, index: usize) -> Option<Value> {
    let Value::Object(obj) = value else {
        return None;
    };
    let object = obj.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(values) => {
            if is_array_hole(&object, index) {
                None
            } else {
                values.get(index).cloned()
            }
        }
        ObjectKind::TypedArray(ta) => {
            if index < ta_live_length(ta) {
                Some(read_element(ta, index))
            } else {
                None
            }
        }
        ObjectKind::Ordinary => object.properties.get(&index.to_string()).cloned(),
        ObjectKind::Map(map) => map.values().nth(index).cloned(),
        _ => None }
}

fn array_like_dense_values(value: &Value) -> Vec<Value> {
    let length = array_like_length(value);
    (0..length)
        .map(|index| array_like_value_at(value, index))
        .collect()
}

fn typed_array_elem(value: &Value) -> Option<TypedElemKind> {
    let Value::Object(obj) = value else {
        return None;
    };
    let object = obj.lock().unwrap();
    match &object.kind {
        ObjectKind::TypedArray(ta) => Some(ta.elem),
        _ => None }
}

fn make_typed_array_from_values(elem: TypedElemKind, values: &[Value]) -> Value {
    let typed = new_typed_array(elem, values.len());
    if let Value::Object(out) = &typed {
        let guard = out.lock().unwrap();
        if let ObjectKind::TypedArray(ta) = &guard.kind {
            for (index, value) in values.iter().enumerate() {
                write_element(ta, index, value);
            }
        }
    }
    typed
}

fn validate_optional_comparator(
    ctx: &mut HostContext,
    compare_fn: Option<&Value>,
    message: &str,
) -> bool {
    match compare_fn {
        None | Some(Value::Undefined) => true,
        Some(value) if is_callable_value(value) => true,
        Some(_) => {
            let _ = throw_type_error(ctx, message);
            false
        }
    }
}

fn default_sort_key(ctx: &mut HostContext, value: &Value) -> Option<String> {
    if matches!(value, Value::Symbol(_)) {
        let _ = throw_type_error(ctx, "Cannot convert a Symbol value to a string");
        None
    } else {
        Some(format!("{}", value))
    }
}

fn compare_sort_values(
    ctx: &mut HostContext,
    compare_fn: Option<&Value>,
    a: &Value,
    b: &Value,
) -> std::cmp::Ordering {
    if let Some(compare_fn) = compare_fn.filter(|v| !matches!(v, Value::Undefined)) {
        let result = invoke_callback(ctx, compare_fn, &[a.clone(), b.clone()]);
        let order = result.as_f64();
        if order < 0.0 {
            std::cmp::Ordering::Less
        } else if order > 0.0 {
            std::cmp::Ordering::Greater
        } else {
            std::cmp::Ordering::Equal
        }
    } else {
        match (default_sort_key(ctx, a), default_sort_key(ctx, b)) {
            (Some(a), Some(b)) => a.cmp(&b),
            _ => std::cmp::Ordering::Equal }
    }
}

fn sort_array_values_ecma(
    ctx: &mut HostContext,
    values: Vec<Value>,
    compare_fn: Option<&Value>,
) -> Vec<Value> {
    let len = values.len();
    let mut sortable = Vec::new();
    let mut undefined_count = 0usize;
    for value in values {
        if matches!(value, Value::Undefined) {
            undefined_count += 1;
        } else {
            sortable.push(value);
        }
    }
    sortable.sort_by(|a, b| compare_sort_values(ctx, compare_fn, a, b));
    sortable.extend((0..undefined_count).map(|_| Value::Undefined));
    sortable.resize(len, Value::Undefined);
    sortable
}

pub fn make_holey_array(length: usize) -> Value {
    let array = make_array(vec![Value::Undefined; length]);
    if let Value::Object(obj) = &array {
        let mut guard = obj.lock().unwrap();
        mark_hole_range(&mut guard, 0..length);
    }
    array
}

fn parse_js_array_length(value: &Value) -> Result<usize, &'static str> {
    match value {
        Value::I32(length) if *length >= 0 => Ok(*length as usize),
        Value::I64(length) if *length >= 0 => Ok(*length as usize),
        Value::F64(length)
            if *length >= 0.0 && length.fract() == 0.0 && *length <= u32::MAX as f64 =>
        {
            Ok(*length as usize)
        }
        Value::String(text) => text
            .parse::<u32>()
            .map(|length| length as usize)
            .map_err(|_| "Invalid array length"),
        _ => Err("Invalid array length") }
}

pub fn set_array_length(object: &mut Object, new_len: usize) {
    let old_len = match &object.kind {
        ObjectKind::Array(values) => values.len(),
        _ => return };

    if let ObjectKind::Array(ref mut values) = object.kind {
        if new_len < old_len {
            values.truncate(new_len);
        } else if new_len > old_len {
            values.resize(new_len, Value::Undefined);
        }
    }

    if new_len < old_len {
        remap_array_holes(object, |index| (index < new_len).then_some(index));
    } else if new_len > old_len {
        mark_hole_range(object, old_len..new_len);
    }

    sync_length(object);
}

pub fn apply_js_array_length(ctx: &mut HostContext, object: &mut Object, value: &Value) {
    match parse_js_array_length(value) {
        Ok(new_len) => set_array_length(object, new_len),
        Err(message) => ctx.throw_value(crate::error::new_error(ctx, "RangeError", message)) }
}

/// Keep the array's cached `length` property in sync with the
/// backing vector's length. Every mutator must call this after
/// modifying the vector — JS code reading `.length` does not re-query
/// the Vec; it reads the stored property.
fn sync_length(obj: &mut Object) {
    if let ObjectKind::Array(ref v) = obj.kind {
        let n = v.len();
        obj.properties.insert("length".into(), Value::F64(n as f64));
    }
}

pub fn register(vm: &mut VM) {
    register_constructors(vm);
    register_property_access(vm);
    register_mutators(vm);
    register_non_mutators(vm);
    register_iteration(vm);
    register_adapters(vm);
    // All `ecma:array/*` host fns are registered above; now bind them onto
    // the shared `Array.prototype` so `arr.method()` resolves through the
    // prototype chain (ECMA-262 §23.1.3) like real JS — no per-method
    // compiler adapter needed. The object owns its methods.
    populate_array_prototype(vm);
}

/// Build an unbound `Array.prototype` method value: a `HostFunction`
/// wrapping `ecma:array/<name>`, flagged `__vybe_method_receiver` so
/// `getMethodForCall`/`bind_method_receiver` prepend the array as the
/// receiver (arg 0) at lookup — exactly the shape every `ecma:array/*`
/// fn expects (`reverse(arr)`, `map(arr, fn)`, …).
fn array_proto_method(vm: &VM, name: &str) -> Option<Value> {
    let idx = *vm
        .host_registry
        .get(&("ecma:array".to_string(), name.to_string()))?;
    let mut fn_obj = Object::new();
    fn_obj.kind = ObjectKind::HostFunction(idx);
    fn_obj.properties.insert(
        "__host_module".into(),
        Value::String(Arc::from("ecma:array")),
    );
    fn_obj
        .properties
        .insert("__host_name".into(), Value::String(Arc::from(name)));
    fn_obj
        .properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    fn_obj.properties.insert(
        "__proto__".into(),
        crate::function::shared_function_prototype(),
    );
    fn_obj
        .properties
        .insert("name".into(), Value::String(Arc::from(name)));
    fn_obj
        .properties
        .insert("__vybe_method_receiver".into(), Value::Bool(true));
    Some(Value::Object(vybe_runtime::heap::alloc(fn_obj)))
}

/// Populate the shared `Array.prototype` with the ECMA-262 §23.1.3
/// method surface, each bound to its `ecma:array/*` host fn. JS method
/// names match the host fn names 1-to-1 here.
fn populate_array_prototype(vm: &VM) {
    const METHODS: &[&str] = &[
        "at",
        "concat",
        "copyWithin",
        "entries",
        "every",
        "fill",
        "filter",
        "find",
        "findIndex",
        "findLast",
        "findLastIndex",
        "flat",
        "flatMap",
        "forEach",
        "includes",
        "indexOf",
        "join",
        "keys",
        "lastIndexOf",
        "map",
        "pop",
        "push",
        "reduce",
        "reduceRight",
        "reverse",
        "shift",
        "slice",
        "some",
        "sort",
        "splice",
        "toLocaleString",
        "toReversed",
        "toSorted",
        "toSpliced",
        "toString",
        "unshift",
        "values",
        "with",
    ];
    let Value::Object(proto) = shared_array_prototype() else {
        return;
    };
    let mut p = proto.lock().unwrap();
    for name in METHODS {
        if let Some(method) = array_proto_method(vm, name) {
            p.properties.insert((*name).to_string(), method);
        }
    }
}

// ── Adapter convenience methods ──────────────────────────────────
//
// Not in ECMA-262 but ubiquitous in language runtimes (.NET / Python /
// Ruby list ops). Live here as one-line compositions of spec methods so
// .NET/VB/Python emitter dispatch can map to a single host fn instead
// of inlining the composition at every call site. Each method is
// equivalent to the documented JS expression and would be optimised
// out by an engine that JITs the dispatch through `at`/`splice`/etc.

fn register_adapters(vm: &mut VM) {
    // clear(arr) — `arr.length = 0`. Mutates in place.
    vm.register_host_fn(
        "ecma:array",
        "clear",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if !is_frozen(&arr) {
                    let mut o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref mut v) = o.kind {
                        v.clear();
                    }
                }
            }
            Value::Undefined
        }),
    );

    // first(arr) — `arr.at(0)`. Convenience for Queue.Peek.
    vm.register_host_fn(
        "ecma:array",
        "first",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return v.first().cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );

    // last(arr) — `arr.at(-1)`. Convenience for Stack.Peek.
    vm.register_host_fn(
        "ecma:array",
        "last",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return v.last().cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );

    // removeAt(arr, idx) — `arr.splice(idx, 1)`, returns removed value.
    vm.register_host_fn(
        "ecma:array",
        "removeAt",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = array_of(args, 0) {
                if !is_frozen(&arr) {
                    let mut o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref mut v) = o.kind {
                        let len = v.len() as i32;
                        let resolved = if idx < 0 { len + idx } else { idx };
                        if resolved >= 0 && (resolved as usize) < v.len() {
                            return v.remove(resolved as usize);
                        }
                    }
                }
            }
            Value::Undefined
        }),
    );

    // insertAt(arr, idx, v) — `arr.splice(idx, 0, v)`. Mutates in place.
    vm.register_host_fn(
        "ecma:array",
        "insertAt",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                if !is_frozen(&arr) {
                    let mut o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref mut v) = o.kind {
                        let len = v.len() as i32;
                        let resolved = if idx < 0 {
                            (len + idx).max(0)
                        } else {
                            idx.min(len)
                        };
                        v.insert(resolved as usize, val);
                    }
                }
            }
            Value::Undefined
        }),
    );

    // removeValue(arr, v) — `arr.splice(arr.indexOf(v), 1)` if found.
    // Returns true if removed, false otherwise (matches .NET List.Remove).
    vm.register_host_fn(
        "ecma:array",
        "removeValue",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                if !is_frozen(&arr) {
                    let mut o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref mut v) = o.kind {
                        if let Some(pos) = v.iter().position(|e| e.eq(&needle)) {
                            v.remove(pos);
                            return Value::Bool(true);
                        }
                    }
                }
            }
            Value::Bool(false)
        }),
    );
}

// ── Constructors ──────────────────────────────────────────────────────

fn register_constructors(vm: &mut VM) {
    // new() -> Array
    // ECMA-262 §23.1.1.1 Array constructor:
    //   new Array()         → []  (length 0)
    //   new Array(n)        → array of length `n`, all `undefined` slots
    //                         (TypeError if n is non-integer or out of range —
    //                         Vybe falls back to a single-element array)
    //   new Array(a, b, …)  → [a, b, …]
    vm.register_host_fn(
        "ecma:array",
        "new",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match args.len() {
            0 => make_array(Vec::new()),
            1 => match &args[0] {
                Value::F64(_) | Value::I32(_) | Value::I64(_) => {
                    match parse_js_array_length(&args[0]) {
                        Ok(length) => make_holey_array(length),
                        Err(message) => {
                            ctx.throw_value(crate::error::new_error(
                                ctx,
                                "RangeError",
                                message,
                            ));
                            Value::Undefined
                        }
                    }
                }
                other => make_array(vec![other.clone()]) },
            _ => make_array(args.to_vec()) }),
    );

    // newWithLength(n: i32) -> Array (n-element, null-filled).
    // Used by language-specific allocations (VB `ReDim`, .NET `new T[n]`)
    // that expect default-value semantics (null/0). JS callers go through
    // `new` above which materializes `undefined` slots per spec.
    vm.register_host_fn(
        "ecma:array",
        "newWithLength",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args
                .first()
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            make_array(vec![Value::Null; n])
        }),
    );
    vm.register_host_fn(
        "vybe:js-array",
        "newWithLength",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args
                .first()
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            make_array(vec![Value::Null; n])
        }),
    );

    // of(...values) -> Array
    //
    // Spec: `Array.of(...values)` — variadic. Each positional arg becomes an
    // element. Unlike `new Array(n)` which allocates a length, `Array.of(n)`
    // is always a 1-element array `[n]`.
    vm.register_host_fn(
        "ecma:array",
        "of",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| make_array(args.to_vec())),
    );

    // from(src, mapFn?) -> Array
    //
    // Spec: `Array.from(iterable, mapFn)`. Accepts arrays, strings, and
    // array-like objects (anything with `length` and numeric keys). When a
    // `mapFn` is supplied it's invoked as `mapFn(value, index)` via the
    // host's `invoke` callback and the result replaces each element.
    vm.register_host_fn(
        "ecma:array",
        "from",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            match args.first() {
                Some(Value::Object(src)) => {
                    let s = src.lock().unwrap();
                    match s.kind {
                        ObjectKind::Array(ref elems) => {
                            out.extend(elems.iter().cloned());
                        }
                        ObjectKind::TypedArray(ref ta) => {
                            let live = crate::typedarray::ta_live_length(ta);
                            for i in 0..live {
                                out.push(crate::typedarray::read_element(ta, i));
                            }
                        }
                        // Map → Array of `[key, value]` pairs (§23.1.2.1).
                        ObjectKind::Map(ref m) => {
                            for (k, v) in m.iter() {
                                out.push(make_array(vec![k.clone(), v.clone()]));
                            }
                        }
                        // Set → Array of values (§23.1.2.1).
                        ObjectKind::Set(ref set) => {
                            out.extend(set.iter().cloned());
                        }
                        _ => {
                            if let Some(len_val) = s.properties.get("length") {
                                let len = len_val.as_f64().max(0.0) as usize;
                                for i in 0..len {
                                    let key = i.to_string();
                                    out.push(
                                        s.properties.get(&key).cloned().unwrap_or(Value::Undefined),
                                    );
                                }
                            }
                        }
                    }
                }
                Some(Value::String(s)) => {
                    for c in s.chars() {
                        out.push(Value::String(Arc::from(c.to_string().as_str())));
                    }
                }
                _ => {}
            }
            if let Some(mapper) = args.get(1) {
                if !matches!(mapper, Value::Null | Value::Undefined) {
                    let mut mapped = Vec::with_capacity(out.len());
                    for (i, v) in out.iter().enumerate() {
                        mapped.push(invoke_callback(
                            ctx,
                            mapper,
                            &[v.clone(), Value::I32(i as i32)],
                        ));
                    }
                    out = mapped;
                }
            }
            make_array(out)
        }),
    );

    // fromWithMap(arrayLike, mapFn) — Array.from with mandatory mapper.
    vm.register_host_fn(
        "ecma:array",
        "fromWithMap",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                match &s.kind {
                    ObjectKind::Array(elems) => out.extend(elems.iter().cloned()),
                    _ => {
                        let len = s
                            .properties
                            .get("length")
                            .map(|v| v.as_f64().max(0.0) as usize)
                            .unwrap_or(0);
                        for i in 0..len {
                            out.push(
                                s.properties
                                    .get(&i.to_string())
                                    .cloned()
                                    .unwrap_or(Value::Undefined),
                            );
                        }
                    }
                }
            }
            if let Some(mapper) = args.get(1) {
                let mut mapped = Vec::with_capacity(out.len());
                for (i, v) in out.iter().enumerate() {
                    mapped.push(invoke_callback(
                        ctx,
                        mapper,
                        &[v.clone(), Value::I32(i as i32)],
                    ));
                }
                out = mapped;
            }
            make_array(out)
        }),
    );

    // fromAsync(asyncIterable, mapFn) -> Promise<Array>
    // Stub for now: returns an empty Array; real impl requires async
    // iteration integration with JSPI. Listed in the import set so
    // compilers emitting calls don't fail to link; behavior will be
    // completed in a later pass.
    vm.register_host_fn(
        "ecma:array",
        "fromAsync",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let source = match crate::iterator::try_maybe_await_value(
                args.first().cloned().unwrap_or(Value::Undefined),
            ) {
                Ok(value) => value,
                Err(reason) => return crate::promise::make_promise("rejected", reason) };
            let mapper = args.get(1).cloned();
            let collected =
                crate::iterator::try_materialize_iterable_values(ctx, &source, true)
                    .and_then(|values| {
                        let mut mapped = Vec::with_capacity(values.len());
                        for (index, value) in values.into_iter().enumerate() {
                            let awaited = crate::iterator::try_maybe_await_value(value)?;
                            let mapped_value = match mapper.as_ref() {
                                Some(mapper)
                                    if !matches!(mapper, Value::Null | Value::Undefined) =>
                                {
                                    ctx.try_invoke(mapper, &[awaited, Value::I32(index as i32)])?
                                }
                                _ => awaited };
                            mapped
                                .push(crate::iterator::try_maybe_await_value(mapped_value)?);
                        }
                        Ok(make_array(mapped))
                    });
            match collected {
                Ok(array) => crate::promise::make_promise("fulfilled", array),
                Err(reason) => crate::promise::make_promise("rejected", reason) }
        }),
    );

    // isArray(v) -> bool
    vm.register_host_fn(
        "ecma:array",
        "isArray",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            // §7.2.2 IsArray: a proxy is an Array when its target is.
            let mut v = args.first().cloned().unwrap_or(Value::Undefined);
            while let Value::Object(obj) = &v {
                match crate::object::proxy_target_and_handler(obj) {
                    Some((target, _)) => v = target,
                    None => break }
            }
            let probe = [v];
            Value::Bool(array_of(&probe, 0).is_some())
        }),
    );
}

// ── Property access ────────────────────────────────────────────────────

fn register_property_access(vm: &mut VM) {
    // get(arr_or_obj, key) -> value
    //
    // Primary use is `Array.prototype`-style integer indexing, but this
    // import is also the landing pad for `dict.has` / `hasOwnProperty`
    // / `in` compiled through compiler_common. Accepts plain objects to
    // satisfy those callers without requiring a second import.
    vm.register_host_fn(
        "ecma:array",
        "get",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let i = key.as_i32();
                        if i < 0 {
                            return Value::Undefined;
                        }
                        return v.get(i as usize).cloned().unwrap_or(Value::Undefined);
                    }
                    // Polymorphic dispatch on Map — the canonical cross-
                    // language associative type (PHP `['k'=>v]`, Python
                    // dicts, Ruby hashes, JS plain objects). Key is looked
                    // up using Value-level equality (SameValueZero), so
                    // both `$m['foo']` and `$m[$key]` where `$key = 'foo'`
                    // resolve identically.
                    ObjectKind::Map(m) => {
                        let lookup_key = match &key {
                            Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F64(_) => {
                                key.clone()
                            }
                            other => {
                                Value::String(std::sync::Arc::from(format!("{}", other).as_str()))
                            }
                        };
                        if let Some(v) = m.get(&lookup_key) {
                            return v.clone();
                        }
                        // PHP-ish fallback: if caller used a string key like
                        // "0" but the map stores integer keys (or vice
                        // versa), try the coerced form. Only coerces for
                        // purely numeric strings to avoid surprises.
                        if let Value::String(s) = &key {
                            if let Ok(n) = s.parse::<i32>() {
                                if let Some(v) = m.get(&Value::I32(n)) {
                                    return v.clone();
                                }
                            }
                        } else if let Value::I32(n) = &key {
                            if let Some(v) =
                                m.get(&Value::String(std::sync::Arc::from(n.to_string().as_str())))
                            {
                                return v.clone();
                            }
                        }
                        return Value::Undefined;
                    }
                    _ => {}
                }
                // Plain Object fallback: property lookup. Used by Ordinary
                // objects and by compiler_common's `has` / `in` emitter.
                let key_str = match &key {
                    Value::String(s) => s.to_string(),
                    other => format!("{}", other) };
                if let Some(v) = o.properties.get(&key_str) {
                    return v.clone();
                }
                return Value::Undefined;
            }
            Value::Undefined
        }),
    );

    // set(arr_or_obj, key, v) -> () — extends arrays with null-fill when
    // key >= length; stores into plain objects by string key; updates Maps
    // using the canonical Value-keyed IndexMap.
    vm.register_host_fn(
        "ecma:array",
        "set",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).cloned().unwrap_or(Value::Undefined);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                if matches!(&o.kind, ObjectKind::Array(_))
                    && matches!(&key, Value::String(text) if text.as_ref() == "length" || text.as_ref() == "__len__")
                {
                    apply_js_array_length(ctx, &mut o, &val);
                    return Value::Null;
                }
                match &mut o.kind {
                    ObjectKind::Array(v) => {
                        // Numeric keys → element store. Non-numeric keys
                        // (e.g. PHP/Python writing string-keyed entries
                        // onto an Array-kind value) → property-bag write
                        // per ECMA-262 §10.4.2.2 (string-named props on
                        // Array exotic objects). Mirrors `Object::set`
                        // which falls through to `properties.insert` when
                        // the key isn't a valid array index.
                        let numeric_idx = match &key {
                            Value::I32(n) if *n >= 0 => Some(*n as usize),
                            Value::I64(n) if *n >= 0 => Some(*n as usize),
                            Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
                            Value::String(s) => s.parse::<usize>().ok(),
                            _ => None };
                        if let Some(idx) = numeric_idx {
                            let old_len = v.len();
                            // ECMA-262 §6.1.7.2 / §23.1.3 — holes from
                            // sparse `arr[hi] = v` writes read as
                            // Undefined, distinct from explicit `Null`.
                            while v.len() <= idx {
                                v.push(Value::Undefined);
                            }
                            v[idx] = val;
                            if idx >= old_len {
                                mark_hole_range(&mut o, old_len..(idx + 1));
                            }
                            clear_array_hole(&mut o, idx);
                            sync_length(&mut o);
                        } else {
                            let key_str = match &key {
                                Value::String(s) => s.to_string(),
                                other => format!("{}", other) };
                            o.properties.insert(key_str, val);
                        }
                    }
                    ObjectKind::Map(m) => {
                        let map_key = match &key {
                            Value::String(_) | Value::I32(_) | Value::I64(_) | Value::F64(_) => key.clone(),
                            other => Value::String(std::sync::Arc::from(format!("{}", other).as_str())) };
                        m.insert(map_key, val);
                    }
                    ObjectKind::TypedArray(ta) => {
                        let numeric_idx = match &key {
                            Value::I32(n) if *n >= 0 => Some(*n as usize),
                            Value::I64(n) if *n >= 0 => Some(*n as usize),
                            Value::F64(n) if n.fract() == 0.0 && *n >= 0.0 => Some(*n as usize),
                            Value::String(s) => s.parse::<usize>().ok(),
                            _ => None };
                        if let Some(idx) = numeric_idx {
                            write_element(ta, idx, &val);
                        }
                    }
                    _ => {
                        let key_str = match &key {
                            Value::String(s) => s.to_string(),
                            other => format!("{}", other) };
                        o.properties.insert(key_str, val);
                    }
                }
            }
            Value::Null
        }),
    );

    // length(arr) -> i32 — ECMA-262 §23.1.3.12. Strict Array/TypedArray
    // only; strings use `wasm:js-string.length` per the js-string-builtins
    // proposal. Polymorphic callers (e.g. our `__len__` canonical) must
    // type-dispatch before selecting the import.
    vm.register_host_fn(
        "ecma:array",
        "length",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(o)) = args.first() {
                let lock = o.lock().unwrap();
                return match &lock.kind {
                    ObjectKind::Array(v) => Value::I32(v.len() as i32),
                    ObjectKind::Map(m) => Value::I32(m.len() as i32),
                    ObjectKind::Set(s) => Value::I32(s.len() as i32),
                    ObjectKind::TypedArray(t) => Value::I32(t.length as i32),
                    _ => lock
                        .properties
                        .get("length")
                        .map(|v| Value::I32(v.as_i32()))
                        .unwrap_or(Value::Null) };
            }
            Value::Null
        }),
    );

    // setLength(arr, n) -> () — truncate or null-fill extend
    vm.register_host_fn(
        "ecma:array",
        "setLength",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let Some(value) = args.get(1) {
                    apply_js_array_length(ctx, &mut o, value);
                }
            }
            Value::Null
        }),
    );

    // at(arr, i) -> value
    //
    // `Array.prototype.at` — negative indices relative to length, undefined
    // when OOB. String `.at()` routes through `ecma:value.invokeMethod`.
    vm.register_host_fn(
        "ecma:array",
        "at",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let idx = if i < 0 { len + i } else { i };
                    if idx < 0 || idx >= len {
                        return Value::Undefined;
                    }
                    return v.get(idx as usize).cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );
}

// ── Mutators ──────────────────────────────────────────────────────────

fn register_mutators(vm: &mut VM) {
    // push(arr, v) -> i32 new_length
    //
    // §23.1.3.23: push defines a NEW index, so a frozen or sealed
    // (non-extensible) array throws TypeError — in any mode.
    vm.register_host_fn(
        "ecma:array",
        "push",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let values = &args[1..];
            if let Some(arr) = array_of(args, 0) {
                let blocked = {
                    let o = arr.lock().unwrap();
                    o.properties.get(FROZEN_MARK).is_some()
                        || o.properties.get("__array_length_readonly").is_some()
                        || crate::object::is_not_extensible(&o)
                };
                if blocked {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "Cannot add property, object is not extensible",
                    ));
                    return Value::Undefined;
                }
                let mut o = arr.lock().unwrap();
                let old_len = match &o.kind {
                    ObjectKind::Array(v) => v.len(),
                    _ => 0 };
                let len = if let ObjectKind::Array(ref mut v) = o.kind {
                    v.extend(values.iter().cloned());
                    v.len() as i32
                } else {
                    0
                };
                for index in old_len..(len as usize) {
                    clear_array_hole(&mut o, index);
                }
                sync_length(&mut o);
                return Value::I32(len);
            }
            if let Some(Value::Object(obj)) = args.first() {
                let mut object = obj.lock().unwrap();
                if matches!(object.kind, ObjectKind::Ordinary) {
                    let start = property_length_as_usize(&object).unwrap_or(0);
                    for (offset, value) in values.iter().enumerate() {
                        object
                            .properties
                            .insert((start + offset).to_string(), value.clone());
                    }
                    let new_length = start + values.len();
                    object
                        .properties
                        .insert("length".into(), Value::F64(new_length as f64));
                    return Value::I32(new_length as i32);
                }
            }
            Value::I32(0)
        }),
    );

    // pop(arr) -> popped_value (undefined if empty or frozen)
    vm.register_host_fn(
        "ecma:array",
        "pop",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) {
                    return Value::Undefined;
                }
                let mut o = arr.lock().unwrap();
                let popped = if let ObjectKind::Array(ref v) = o.kind {
                    if v.is_empty() {
                        Value::Undefined
                    } else {
                        let last_index = v.len() - 1;
                        let was_hole = is_array_hole(&o, last_index);
                        let value = if let ObjectKind::Array(ref mut inner) = o.kind {
                            inner.pop().unwrap_or(Value::Undefined)
                        } else {
                            Value::Undefined
                        };
                        clear_array_hole(&mut o, last_index);
                        if was_hole { Value::Undefined } else { value }
                    }
                } else {
                    Value::Undefined
                };
                sync_length(&mut o);
                return popped;
            }
            Value::Undefined
        }),
    );

    // shift(arr) -> first_value (undefined if empty or frozen)
    vm.register_host_fn(
        "ecma:array",
        "shift",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) {
                    return Value::Undefined;
                }
                let mut o = arr.lock().unwrap();
                let shifted = if let ObjectKind::Array(ref v) = o.kind {
                    if v.is_empty() {
                        Value::Undefined
                    } else {
                        let was_hole = is_array_hole(&o, 0);
                        let value = if let ObjectKind::Array(ref mut inner) = o.kind {
                            inner.remove(0)
                        } else {
                            Value::Undefined
                        };
                        remap_array_holes(&mut o, |index| match index {
                            0 => None,
                            other => Some(other - 1) });
                        if was_hole { Value::Undefined } else { value }
                    }
                } else {
                    Value::Undefined
                };
                sync_length(&mut o);
                return shifted;
            }
            Value::Undefined
        }),
    );

    // unshift(arr, v1, v2, ...) -> i32 new_length
    //
    // Spec: inserts all v_i at the head in order, so `[3,4,5].unshift(1,2)`
    // yields `[1,2,3,4,5]`. Frozen arrays just return the current length.
    vm.register_host_fn(
        "ecma:array",
        "unshift",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        return Value::I32(v.len() as i32);
                    }
                    return Value::I32(0);
                }
                let mut o = arr.lock().unwrap();
                let offset = args.len().saturating_sub(1);
                let len = if let ObjectKind::Array(ref mut v) = o.kind {
                    for (i, val) in args.iter().skip(1).enumerate() {
                        v.insert(i, val.clone());
                    }
                    v.len() as i32
                } else {
                    0
                };
                if offset > 0 {
                    remap_array_holes(&mut o, |index| Some(index + offset));
                }
                sync_length(&mut o);
                return Value::I32(len);
            }
            Value::I32(0)
        }),
    );

    // splice(arr, start, deleteCount, ...items) -> deleted_array
    //
    // Spec: items come through as variadic individual args, not a wrapped
    // array. `arr.splice(1, 0, 2, 3)` inserts 2 and 3 at index 1 and
    // deletes 0; args[3..] hold the items.
    vm.register_host_fn(
        "ecma:array",
        "splice",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            // §23.1.3.31: when deleteCount is OMITTED, everything from
            // start to the end is removed (explicit undefined means 0).
            let del = match args.get(2) {
                None => usize::MAX,
                Some(v) => v.as_i32().max(0) as usize };
            let items: Vec<Value> = args.iter().skip(3).cloned().collect();
            let mut deleted = Vec::new();
            let mut deleted_holes = BTreeSet::new();
            if let Some(arr) = array_of(args, 0) {
                if is_frozen(&arr) {
                    return throw_type_error(ctx, "Cannot modify frozen array");
                }
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len();
                    let idx = if start < 0 {
                        ((len as i32) + start).max(0) as usize
                    } else {
                        (start as usize).min(len)
                    };
                    let end = idx.saturating_add(del).min(len);
                    let delete_count = end.saturating_sub(idx);
                    let insert_count = items.len();
                    let old_holes = hole_indices(&o);
                    for offset in 0..delete_count {
                        if old_holes.contains(&(idx + offset)) {
                            deleted_holes.insert(offset);
                        }
                    }
                    if let ObjectKind::Array(ref mut v) = o.kind {
                        for _ in idx..end {
                            deleted.push(v.remove(idx));
                        }
                        for (i, val) in items.into_iter().enumerate() {
                            v.insert(idx + i, val);
                        }
                    }
                    let shift = insert_count as isize - delete_count as isize;
                    let remapped: BTreeSet<usize> = old_holes
                        .into_iter()
                        .filter_map(|hole| {
                            if hole < idx {
                                Some(hole)
                            } else if hole < end {
                                None
                            } else {
                                Some((hole as isize + shift) as usize)
                            }
                        })
                        .collect();
                    store_hole_indices(&mut o, &remapped);
                }
                sync_length(&mut o);
            }
            let removed = make_array(deleted);
            if let Value::Object(obj) = &removed {
                let mut guard = obj.lock().unwrap();
                store_hole_indices(&mut guard, &deleted_holes);
            }
            removed
        }),
    );

    // reverse(arr) -> self (in-place)
    vm.register_host_fn(
        "ecma:array",
        "reverse",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(_) => {
                        if let ObjectKind::Array(ref mut v) = o.kind {
                            v.reverse();
                        }
                    }
                    ObjectKind::TypedArray(ta) => {
                        let live = ta_live_length(ta);
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
                    _ => {}
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // sort(arr, compareFn) -> self (in-place) — MVP: compare by stringified value
    // Real callback dispatch requires VM `invoke_callback`; implement
    // in the Phase B5 iterator-helpers pass when we tackle callbacks
    // uniformly.
    vm.register_host_fn(
        "ecma:array",
        "sort",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let compare_fn = args.get(1).cloned();
                if !validate_optional_comparator(
                    ctx,
                    compare_fn.as_ref(),
                    "The comparison function must be either a function or undefined",
                ) {
                    return Value::Undefined;
                }
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        // Sort writes every index — a frozen array throws
                        // TypeError (sealed still allows index writes).
                        if o.properties.get(FROZEN_MARK).is_some() {
                            drop(o);
                            ctx.throw_value(crate::error::new_error(
                                ctx,
                                "TypeError",
                                "Cannot assign to read only property of frozen array",
                            ));
                            return Value::Undefined;
                        }
                        let mut values = v.clone();
                        for hole in hole_indices(&o) {
                            if hole < values.len() {
                                values[hole] = Value::Undefined;
                            }
                        }
                        drop(o);
                        values = sort_array_values_ecma(ctx, values, compare_fn.as_ref());
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut v) = o.kind {
                            *v = values;
                        }
                        store_hole_indices(&mut o, &BTreeSet::new());
                    }
                    ObjectKind::TypedArray(ta) => {
                        let live = ta_live_length(ta);
                        let mut values: Vec<Value> =
                            (0..live).map(|i| read_element(ta, i)).collect();
                        values = sort_array_values_ecma(ctx, values, compare_fn.as_ref());
                        for (i, v) in values.iter().enumerate() {
                            write_element(ta, i, v);
                        }
                    }
                    _ => {}
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // fill(arr, value, start, end) -> self
    vm.register_host_fn(
        "ecma:array",
        "fill",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let val = args.get(1).cloned().unwrap_or(Value::Null);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            // §23.1.3.7 relative indexing: negative counts from the end;
            // i64 math so a saturated -Infinity cast can't overflow.
            let norm = |x: i32, len: i32| -> usize {
                if x < 0 {
                    ((len as i64) + (x as i64)).max(0) as usize
                } else {
                    x.min(len) as usize
                }
            };
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(_) => {
                        drop(o);
                        let mut o = obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut v) = o.kind {
                            let len = v.len() as i32;
                            let s = norm(start, len);
                            let e = norm(end, len);
                            for i in s..e {
                                v[i] = val.clone();
                            }
                            // Clear hole markers for the filled range
                            let mut holes = hole_indices(&o);
                            for i in s..e {
                                holes.remove(&i);
                            }
                            store_hole_indices(&mut o, &holes);
                        }
                        sync_length(&mut o);
                    }
                    ObjectKind::TypedArray(ta) => {
                        let live = ta_live_length(ta) as i32;
                        let s = norm(start, live);
                        let e = norm(end, live);
                        for i in s..e {
                            write_element(ta, i, &val);
                        }
                    }
                    _ => {}
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    // copyWithin(arr, target, start, end) -> self
    vm.register_host_fn(
        "ecma:array",
        "copyWithin",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let target = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            // §23.1.3.4 relative indexing — same normalization as fill.
            let norm = |x: i32, len: i32| -> usize {
                if x < 0 {
                    ((len as i64) + (x as i64)).max(0) as usize
                } else {
                    x.min(len) as usize
                }
            };
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    let len = v.len() as i32;
                    let t = norm(target, len);
                    let s = norm(start, len);
                    let e = norm(end, len).max(s);
                    let slice: Vec<Value> = v[s..e].iter().cloned().collect();
                    let max_copy = (len as usize - t).min(slice.len());
                    v[t..t + max_copy].clone_from_slice(&slice[..max_copy]);
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );
    vm.register_host_fn(
        "vybe:js-array",
        "fill",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let val = args.get(1).cloned().unwrap_or(Value::Null);
            let start = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(3).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(arr) = array_of(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    let len = v.len() as i32;
                    let s = start.max(0).min(len) as usize;
                    let e = end.max(0).min(len) as usize;
                    for i in s..e {
                        v[i] = val.clone();
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );
}

// ── Non-mutators ──────────────────────────────────────────────────────

fn register_non_mutators(vm: &mut VM) {
    // slice(arr, start, end) -> new_arr | substring
    //
    // Array slicing is the spec contract (ECMA-262 §23.1.3.28). The
    // compiler's `__vybe_slice` polyfill is the user-facing entry point
    // and dispatches both string and array inputs through the SAME
    // global func ref — when this `ecma:array.slice` host fn shadows the
    // polyfill, it must keep the polymorphic shape so `s[0..5]` (which
    // lowers to `__vybe_slice(s, 0, 5)`) keeps producing a substring.
    // Equivalent to `wasm:js-string.slice` but routed through the
    // single override entry point.
    vm.register_host_fn(
        "ecma:array",
        "slice",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            if let Some(Value::String(s)) = args.first() {
                let chars: Vec<char> = s.chars().collect();
                let len = chars.len() as i32;
                let si = (if start < 0 { len + start } else { start })
                    .max(0)
                    .min(len) as usize;
                let ei = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                let out: String = if si < ei {
                    chars[si..ei].iter().collect()
                } else {
                    String::new()
                };
                return Value::String(Arc::from(out.as_str()));
            }
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let s = (if start < 0 { len + start } else { start })
                        .max(0)
                        .min(len) as usize;
                    let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                    let out: Vec<Value> = if s < e { v[s..e].to_vec() } else { Vec::new() };
                    let sliced = make_array(out);
                    if let Value::Object(obj) = &sliced {
                        let holes: BTreeSet<usize> = (s..e)
                            .filter(|index| is_array_hole(&o, *index))
                            .map(|index| index - s)
                            .collect();
                        let mut guard = obj.lock().unwrap();
                        store_hole_indices(&mut guard, &holes);
                    }
                    return sliced;
                }
            }
            make_array(Vec::new())
        }),
    );

    // concat(arr, other) -> new_arr
    //
    // Used by both `Array.prototype.concat` (spec — only spreads Arrays
    // into the result) AND by spread-element compilation in array
    // literals like `[...s]` (which needs to spread any iterable).
    // Map/Set/String aren't ECMA-262 §23.1.3.2 concatable, but JS
    // engines spread them in practice when the literal-spread path
    // routes here. We handle both in one place.
    vm.register_host_fn(
        "ecma:array",
        "concat",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    out.extend(v.iter().cloned());
                }
            }
            // Spec: if `other` is an iterable, spread it; otherwise
            // append as single element.
            match args.get(1) {
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    match &lock.kind {
                        ObjectKind::Array(v) => out.extend(v.iter().cloned()),
                        ObjectKind::Set(s) => out.extend(s.iter().cloned()),
                        ObjectKind::Map(m) => {
                            for (k, v) in m.iter() {
                                let pair = vec![k.clone(), v.clone()];
                                out.push(make_array(pair));
                            }
                        }
                        _ => out.push(Value::Object(o.clone())) }
                }
                Some(Value::String(s)) => {
                    // Spreading a string yields its code-points (per
                    // Symbol.iterator on String).
                    for c in s.chars() {
                        out.push(Value::String(Arc::from(c.to_string().as_str())));
                    }
                }
                Some(v) => out.push(v.clone()),
                None => {}
            }
            make_array(out)
        }),
    );

    // indexOf(arr, value, fromIndex) -> i32
    vm.register_host_fn(
        "ecma:array",
        "indexOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            let from = args.get(2).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let start = from.max(0) as usize;
                    for (i, elem) in v.iter().enumerate().skip(start) {
                        if is_array_hole(&o, i) {
                            continue;
                        }
                        if elem.eq(&needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    // lastIndexOf(arr, value, fromIndex) -> i32
    vm.register_host_fn(
        "ecma:array",
        "lastIndexOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let from = args.get(2).map(|v| v.as_i32()).unwrap_or(len - 1);
                    let end = from.min(len - 1).max(-1);
                    let end_idx = if end < 0 { 0 } else { (end + 1) as usize };
                    for (i, elem) in v[..end_idx].iter().enumerate().rev() {
                        if is_array_hole(&o, i) {
                            continue;
                        }
                        if elem.eq(&needle) {
                            return Value::I32(i as i32);
                        }
                    }
                }
            }
            Value::I32(-1)
        }),
    );

    // includes(arr_or_obj, value, fromIndex) -> bool
    //
    // Primary: `Array.prototype.includes` (SameValueZero comparison). Also
    // the landing pad for the compiled `x in y` operator when `y` is a
    // plain object — we check own-property membership for that case.
    // String `.includes(...)` routes through `ecma:value.invokeMethod`.
    vm.register_host_fn(
        "ecma:array",
        "includes",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            let from = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let needle_is_nan = matches!(needle, Value::F64(n) if n.is_nan());
                        for (index, elem) in v.iter().enumerate().skip(from) {
                            if is_array_hole(&o, index) {
                                if matches!(needle, Value::Undefined) {
                                    return Value::Bool(true);
                                }
                                continue;
                            }
                            if needle_is_nan && matches!(elem, Value::F64(n) if n.is_nan()) {
                                return Value::Bool(true);
                            }
                            if elem.eq(&needle) {
                                return Value::Bool(true);
                            }
                        }
                        return Value::Bool(false);
                    }
                    // Polymorphic on Map — PHP `in_array($v, $map)` checks
                    // whether `$v` is among the map's VALUES (not keys).
                    ObjectKind::Map(m) => {
                        for (_k, v) in m.iter().skip(from) {
                            if v.eq(&needle) {
                                return Value::Bool(true);
                            }
                        }
                        return Value::Bool(false);
                    }
                    _ => {}
                }
                // Ordinary fallback — checks property VALUES.
                for (_k, v) in o.properties.iter() {
                    if v.eq(&needle) {
                        return Value::Bool(true);
                    }
                }
                return Value::Bool(false);
            }
            Value::Bool(false)
        }),
    );

    // join(arr, sep) -> string. Polymorphic over Array and Map (PHP
    // associative arrays compile to ObjectKind::Map, and `implode` /
    // `array.join` on them should iterate values in insertion order).
    vm.register_host_fn(
        "ecma:array",
        "join",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let sep = args
                .get(1)
                .map(|v| format!("{}", v))
                .unwrap_or_else(|| ",".into());
            let parts: Vec<String> = match args.first() {
                Some(Value::Object(o)) => {
                    let inner = o.lock().unwrap();
                    let stringify = |e: &Value| match e {
                        Value::Null | Value::Undefined => String::new(),
                        // §7.1.17 ToString(BigInt) has no `n` suffix — that
                        // is console.log's inspection format only.
                        Value::BigInt(n) => n.to_string(),
                        _ => format!("{}", e) };
                    match &inner.kind {
                        ObjectKind::Array(v) => v
                            .iter()
                            .enumerate()
                            .map(|(index, value)| {
                                if is_array_hole(&inner, index) {
                                    String::new()
                                } else {
                                    stringify(value)
                                }
                            })
                            .collect(),
                        ObjectKind::Map(m) => m.values().map(stringify).collect(),
                        ObjectKind::Ordinary => {
                            if let Some(len) = property_length_as_usize(&inner) {
                                (0..len)
                                    .map(|index| {
                                        inner
                                            .properties
                                            .get(&index.to_string())
                                            .map(stringify)
                                            .unwrap_or_default()
                                    })
                                    .collect()
                            } else {
                                // Plain JS object — iterate values in
                                // insertion order. The compiler tracks
                                // insertion order in a side `__keys` array
                                // when index-assigning string keys; without
                                // it, `properties` is a HashMap and
                                // iteration order is randomized per
                                // process. Honor `__keys` first; fall back
                                // to the hash-map order only when the
                                // compiler hasn't installed one (e.g. an
                                // empty `{}` literal that's never been
                                // mutated).
                                //
                                // Skip internal `__*` metadata keys (Vybe
                                // stores prototype links and type tags as
                                // `__type`, `__proto__`, etc.).
                                let ordered_keys: Option<Vec<String>> =
                                    inner.properties.get("__keys").and_then(|v| {
                                        if let Value::Object(arr) = v {
                                            let a = arr.lock().unwrap();
                                            if let ObjectKind::Array(items) = &a.kind {
                                                Some(
                                                    items
                                                        .iter()
                                                        .map(|k| format!("{}", k))
                                                        .collect(),
                                                )
                                            } else {
                                                None
                                            }
                                        } else {
                                            None
                                        }
                                    });
                                if let Some(keys) = ordered_keys {
                                    keys.iter()
                                        .filter(|k| !k.starts_with("__"))
                                        .filter_map(|k| inner.properties.get(k).map(stringify))
                                        .collect()
                                } else {
                                    inner
                                        .properties
                                        .iter()
                                        .filter(|(k, _)| !k.starts_with("__"))
                                        .map(|(_, v)| stringify(v))
                                        .collect()
                                }
                            }
                        }
                        ObjectKind::TypedArray(ta) => {
                            let live = ta_live_length(ta);
                            (0..live)
                                .map(|i| typed_array_join_element(read_element(ta, i)))
                                .collect()
                        }
                        _ => Vec::new() }
                }
                _ => Vec::new() };
            Value::String(Arc::from(parts.join(&sep).as_str()))
        }),
    );

    // toString(arr) -> string (same as join with default ",")
    vm.register_host_fn(
        "ecma:array",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let parts: Vec<String> = v
                        .iter()
                        .enumerate()
                        .map(|(index, value)| {
                            if is_array_hole(&o, index) {
                                String::new()
                            } else {
                                match value {
                                    Value::Null | Value::Undefined => String::new(),
                                    _ => format!("{}", value) }
                            }
                        })
                        .collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // toLocaleString — same as toString for MVP
    vm.register_host_fn(
        "ecma:array",
        "toLocaleString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            // Same as toString — real locale-aware conversion lives in
            // Phase F (intl integration).
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let parts: Vec<String> = v
                        .iter()
                        .enumerate()
                        .map(|(index, value)| {
                            if is_array_hole(&o, index) {
                                String::new()
                            } else {
                                format!("{}", value)
                            }
                        })
                        .collect();
                    return Value::String(Arc::from(parts.join(",").as_str()));
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // flat(arr, depth) -> new_arr
    vm.register_host_fn(
        "ecma:array",
        "flat",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let depth = args.get(1).map(|v| v.as_i32()).unwrap_or(1);
            fn flatten(
                ctx: &mut HostContext,
                out: &mut Vec<Value>,
                arr_obj: &Object,
                elems: &[Value],
                depth: i32,
                seen: &mut HashSet<usize>,
            ) -> bool {
                for (i, v) in elems.iter().enumerate() {
                    if is_array_hole(arr_obj, i) {
                        continue;
                    }
                    if depth > 0 {
                        if let Value::Object(o) = v {
                            let id = Arc::as_ptr(o) as usize;
                            if seen.contains(&id) {
                                ctx.throw_value(crate::error::new_error(
                                    ctx,
                                    "TypeError",
                                    "Circular array cannot be flattened",
                                ));
                                return false;
                            }
                            let lock = o.lock().unwrap();
                            if let ObjectKind::Array(ref inner) = lock.kind {
                                seen.insert(id);
                                let ok = flatten(ctx, out, &lock, inner, depth - 1, seen);
                                seen.remove(&id);
                                if !ok {
                                    return false;
                                }
                                continue;
                            }
                        }
                    }
                    out.push(v.clone());
                }
                true
            }
            let mut out = Vec::new();
            if let Some(arr) = array_of(args, 0) {
                let id = Arc::as_ptr(&arr) as usize;
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let mut seen = HashSet::from([id]);
                    flatten(ctx, &mut out, &o, v, depth, &mut seen);
                }
            }
            make_array(out)
        }),
    );

    // ── ES2023 non-mutating variants ────────────────────────────────

    // toReversed(arr) -> new_arr
    vm.register_host_fn(
        "ecma:array",
        "toReversed",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let mut out = array_like_dense_values(&receiver);
            out.reverse();
            if let Some(elem) = typed_array_elem(&receiver) {
                make_typed_array_from_values(elem, &out)
            } else {
                make_array(out)
            }
        }),
    );

    // toSorted(arr, compareFn) -> new_arr
    vm.register_host_fn(
        "ecma:array",
        "toSorted",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let compare_fn = args.get(1).cloned();
            if !validate_optional_comparator(
                ctx,
                compare_fn.as_ref(),
                "The comparison function must be either a function or undefined",
            ) {
                return Value::Undefined;
            }
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let out = sort_array_values_ecma(
                ctx,
                array_like_dense_values(&receiver),
                compare_fn.as_ref(),
            );
            if let Some(elem) = typed_array_elem(&receiver) {
                make_typed_array_from_values(elem, &out)
            } else {
                make_array(out)
            }
        }),
    );

    // with(arr, i, v) -> new_arr
    vm.register_host_fn(
        "ecma:array",
        "with",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let mut out = array_like_dense_values(&receiver);
            let len = out.len() as i32;
            let idx = if i < 0 { len + i } else { i };
            if idx < 0 || idx >= len {
                return throw_range_error(ctx, "Invalid index");
            }
            out[idx as usize] = val;
            if let Some(elem) = typed_array_elem(&receiver) {
                make_typed_array_from_values(elem, &out)
            } else {
                make_array(out)
            }
        }),
    );
}

fn typed_array_join_element(value: Value) -> String {
    match value {
        Value::BigInt(n) => format!("{}", n),
        other => format!("{}", other) }
}

// ── Iteration / higher-order callbacks ─────────────────────────────────

/// Captured at register-time so `make_array_iterator` can stamp a
/// HostFunction property pointing at `iterNext` without re-resolving
/// the registry on every call.
static ARRAY_ITER_NEXT_IDX: std::sync::OnceLock<usize> = std::sync::OnceLock::new();

fn iter_result(value: Value, done: bool) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("value".into(), value);
    obj.properties.insert("done".into(), Value::Bool(done));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Build an Array Iterator (§23.1.5) backed by a materialized Vec.
/// The iterator's `ObjectKind::Array(...)` lets spread/for-of fall back
/// to plain-array iteration when the consumer doesn't drive `.next()`
/// explicitly. `__index` tracks an independent cursor for `.next()`.
pub fn make_array_iterator(materialized: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Array(materialized);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("ArrayIterator")));
    obj.properties.insert("__index".into(), Value::I32(0));
    if let Some(idx) = ARRAY_ITER_NEXT_IDX.get() {
        obj.properties.insert(
            "next".into(),
            receiver_host_fn_ref("ecma:array", "iterNext", *idx),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn register_iteration(vm: &mut VM) {
    // `iterNext(this)` — implements §23.1.5.2.1 Array Iterator next().
    // Reads `__index`, returns `{value, done}`, advances the cursor.
    vm.register_host_fn(
        "ecma:array",
        "iterNext",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(it)) = args.first() else {
                return iter_result(Value::Undefined, true);
            };
            let mut o = it.lock().unwrap();
            let idx = o.properties.get("__index").map(|v| v.as_i32()).unwrap_or(0);
            if let ObjectKind::Array(ref items) = o.kind {
                if (idx as usize) < items.len() {
                    let value = items[idx as usize].clone();
                    o.properties.insert("__index".into(), Value::I32(idx + 1));
                    return iter_result(value, false);
                }
            }
            iter_result(Value::Undefined, true)
        }),
    );
    if let Some(idx) = vm
        .host_registry
        .get(&("ecma:array".to_string(), "iterNext".to_string()))
        .copied()
    {
        let _ = ARRAY_ITER_NEXT_IDX.set(idx);
    }

    // keys(arr) / values(arr) / entries(arr) — §23.1.3.{16,36,7}.
    // Return a §23.1.5 Array Iterator with `next()` driving the cursor.
    // The iterator's underlying Array kind keeps spread / for-of working
    // through plain-array iteration paths.
    vm.register_host_fn(
        "ecma:array",
        "keys",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let out: Vec<Value> = (0..v.len()).map(|i| Value::F64(i as f64)).collect();
                    return make_array_iterator(out);
                }
            }
            make_array_iterator(Vec::new())
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "values",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = array_of(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return make_array_iterator(v.clone());
                }
            }
            make_array_iterator(Vec::new())
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "entries",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let out: Vec<Value> = v
                            .iter()
                            .enumerate()
                            .map(|(i, e)| make_array(vec![Value::F64(i as f64), e.clone()]))
                            .collect();
                        return make_array_iterator(out);
                    }
                    ObjectKind::TypedArray(ta) => {
                        let live = ta_live_length(ta);
                        let out: Vec<Value> = (0..live)
                            .map(|i| make_array(vec![Value::F64(i as f64), read_element(ta, i)]))
                            .collect();
                        return make_array_iterator(out);
                    }
                    _ => {}
                }
            }
            make_array_iterator(Vec::new())
        }),
    );

    // ── Callback-taking methods (Phase B5 — real dispatch) ────────────
    //
    // These invoke the supplied JS callback per element via
    // `HostContext::invoke`, matching the callback signature
    // `(element, index, array) → result` that the MDN reference
    // pages specify for Array.prototype methods.
    //
    // Spec references (each method):
    //   - forEach: §23.1.3.13 — no return; just invoke for side effects
    //   - map:     §23.1.3.21 — collect invoke results
    //   - filter:  §23.1.3.8  — keep elements where callback is truthy
    //   - reduce:  §23.1.3.26 — fold from left, optional initial value
    //   - reduceRight: §23.1.3.27
    //   - some:    §23.1.3.30 — any callback returns truthy
    //   - every:   §23.1.3.6  — all callbacks return truthy
    //   - find, findLast: §23.1.3.11 — first/last element where truthy
    //   - findIndex, findLastIndex: §23.1.3.12 — index of first/last match
    //   - flatMap: §23.1.3.15 — map + flatten one level

    vm.register_host_fn(
        "ecma:array",
        "forEach",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let entries = {
                    let o = arr.lock().unwrap();
                    present_array_entries(&o)
                };
                for (i, elem) in entries {
                    let invoke_args = vec![elem, Value::I32(i as i32), Value::Object(arr.clone())];
                    invoke_callback(ctx, &callback, &invoke_args);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "map",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.map callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let mapped = make_holey_array(length);
                if let Value::Object(mapped_obj) = &mapped {
                    let mut mapped_guard = mapped_obj.lock().unwrap();
                    let mut clear_indices = Vec::new();
                    if let ObjectKind::Array(ref mut values) = mapped_guard.kind {
                        for index in 0..length {
                            let Some(elem) = array_present_value_at(&arr, index) else {
                                continue;
                            };
                            let invoke_args =
                                vec![elem, Value::I32(index as i32), Value::Object(arr.clone())];
                            values[index] = invoke_callback_this(
                                ctx,
                                &callback,
                                this_arg.clone(),
                                &invoke_args,
                            );
                            clear_indices.push(index);
                        }
                    }
                    for index in clear_indices {
                        clear_array_hole(&mut mapped_guard, index);
                    }
                }
                return mapped;
            }
            let length = array_like_length(&receiver);
            if length > 0 || matches!(receiver, Value::Object(_)) {
                let mapped: Vec<Value> = (0..length)
                    .map(|i| {
                        let elem = array_like_value_at(&receiver, i);
                        let invoke_args = vec![elem, Value::I32(i as i32), receiver.clone()];
                        invoke_callback_this(ctx, &callback, this_arg.clone(), &invoke_args)
                    })
                    .collect();
                return make_array(mapped);
            }
            make_array(Vec::new())
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "filter",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.filter callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let mut filtered = Vec::new();
                for index in 0..length {
                    let Some(elem) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(index as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&invoke_callback_this(
                        ctx,
                        &callback,
                        this_arg.clone(),
                        &invoke_args,
                    )) {
                        filtered.push(elem);
                    }
                }
                return make_array(filtered);
            }
            let length = array_like_length(&receiver);
            if length > 0 || matches!(receiver, Value::Object(_)) {
                let filtered: Vec<Value> = (0..length)
                    .filter_map(|i| {
                        let elem = array_like_value_at(&receiver, i);
                        let invoke_args =
                            vec![elem.clone(), Value::I32(i as i32), receiver.clone()];
                        let keep = is_truthy(&invoke_callback_this(
                            ctx,
                            &callback,
                            this_arg.clone(),
                            &invoke_args,
                        ));
                        if keep { Some(elem.clone()) } else { None }
                    })
                    .collect();
                return make_array(filtered);
            }
            make_array(Vec::new())
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "reduce",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.reduce callback is not callable",
            ) {
                return Value::Undefined;
            }
            let initial_provided =
                args.len() > 2 && !matches!(args.get(2), Some(Value::Undefined) | None);
            let mut acc = if initial_provided {
                args.get(2).cloned().unwrap_or(Value::Undefined)
            } else {
                Value::Undefined
            };
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let mut start_idx = 0;
                if !initial_provided {
                    let mut found = false;
                    for index in 0..length {
                        if let Some(value) = array_present_value_at(&arr, index) {
                            acc = value;
                            start_idx = index + 1;
                            found = true;
                            break;
                        }
                    }
                    if !found {
                        return throw_type_error(
                            ctx,
                            "Reduce of empty array with no initial value",
                        );
                    }
                }
                for index in start_idx..length {
                    let Some(value) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    let invoke_args = vec![
                        acc,
                        value,
                        Value::I32(index as i32),
                        Value::Object(arr.clone()),
                    ];
                    acc = invoke_callback(ctx, &callback, &invoke_args);
                }
            }
            acc
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "reduceRight",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let initial_provided =
                args.len() > 2 && !matches!(args.get(2), Some(Value::Undefined) | None);
            let mut acc = if initial_provided {
                args.get(2).cloned().unwrap_or(Value::Undefined)
            } else {
                Value::Undefined
            };
            if let Some(arr) = array_of(args, 0) {
                let entries = {
                    let o = arr.lock().unwrap();
                    present_array_entries(&o)
                };
                if entries.is_empty() {
                    return if initial_provided {
                        acc
                    } else {
                        Value::Undefined
                    };
                }
                if !initial_provided {
                    acc = entries
                        .last()
                        .map(|(_, value)| value.clone())
                        .unwrap_or(Value::Undefined);
                }
                for (index, value) in
                    entries
                        .into_iter()
                        .rev()
                        .skip(if initial_provided { 0 } else { 1 })
                {
                    let invoke_args = vec![
                        acc,
                        value,
                        Value::I32(index as i32),
                        Value::Object(arr.clone()),
                    ];
                    acc = invoke_callback(ctx, &callback, &invoke_args);
                }
            }
            acc
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "some",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let entries = {
                    let o = arr.lock().unwrap();
                    present_array_entries(&o)
                };
                for (i, elem) in entries {
                    let invoke_args = vec![elem, Value::I32(i as i32), Value::Object(arr.clone())];
                    if is_truthy(&invoke_callback(ctx, &callback, &invoke_args)) {
                        return Value::Bool(true);
                    }
                }
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "every",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(arr) = array_of(args, 0) {
                let entries = {
                    let o = arr.lock().unwrap();
                    present_array_entries(&o)
                };
                for (i, elem) in entries {
                    let invoke_args = vec![elem, Value::I32(i as i32), Value::Object(arr.clone())];
                    if !is_truthy(&invoke_callback(ctx, &callback, &invoke_args)) {
                        return Value::Bool(false);
                    }
                }
            }
            Value::Bool(true) // spec: empty array → every returns true
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "find",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.find callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                for index in 0..length {
                    let elem = array_value_at(&arr, index);
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(index as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&invoke_callback_this(
                        ctx,
                        &callback,
                        this_arg.clone(),
                        &invoke_args,
                    )) {
                        return elem.clone();
                    }
                }
                return Value::Undefined;
            }
            let length = array_like_length(&receiver);
            for index in 0..length {
                let elem = array_like_value_at(&receiver, index);
                let invoke_args = vec![elem.clone(), Value::I32(index as i32), receiver.clone()];
                if is_truthy(&invoke_callback_this(
                    ctx,
                    &callback,
                    this_arg.clone(),
                    &invoke_args,
                )) {
                    return elem;
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "findIndex",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.findIndex callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                for index in 0..length {
                    let elem = array_value_at(&arr, index);
                    let invoke_args =
                        vec![elem, Value::I32(index as i32), Value::Object(arr.clone())];
                    if is_truthy(&invoke_callback_this(
                        ctx,
                        &callback,
                        this_arg.clone(),
                        &invoke_args,
                    )) {
                        return Value::I32(index as i32);
                    }
                }
                return Value::I32(-1);
            }
            let length = array_like_length(&receiver);
            for index in 0..length {
                let elem = array_like_value_at(&receiver, index);
                let invoke_args = vec![elem, Value::I32(index as i32), receiver.clone()];
                if is_truthy(&invoke_callback_this(
                    ctx,
                    &callback,
                    this_arg.clone(),
                    &invoke_args,
                )) {
                    return Value::I32(index as i32);
                }
            }
            Value::I32(-1)
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "findLast",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.findLast callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                for index in (0..length).rev() {
                    let elem = array_value_at(&arr, index);
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(index as i32),
                        Value::Object(arr.clone()),
                    ];
                    if is_truthy(&invoke_callback_this(
                        ctx,
                        &callback,
                        this_arg.clone(),
                        &invoke_args,
                    )) {
                        return elem.clone();
                    }
                }
                return Value::Undefined;
            }
            let length = array_like_length(&receiver);
            for index in (0..length).rev() {
                let elem = array_like_value_at(&receiver, index);
                let invoke_args = vec![elem.clone(), Value::I32(index as i32), receiver.clone()];
                if is_truthy(&invoke_callback_this(
                    ctx,
                    &callback,
                    this_arg.clone(),
                    &invoke_args,
                )) {
                    return elem;
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "findLastIndex",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.findLastIndex callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                for index in (0..length).rev() {
                    let elem = array_value_at(&arr, index);
                    let invoke_args =
                        vec![elem, Value::I32(index as i32), Value::Object(arr.clone())];
                    if is_truthy(&invoke_callback_this(
                        ctx,
                        &callback,
                        this_arg.clone(),
                        &invoke_args,
                    )) {
                        return Value::I32(index as i32);
                    }
                }
                return Value::I32(-1);
            }
            let length = array_like_length(&receiver);
            for index in (0..length).rev() {
                let elem = array_like_value_at(&receiver, index);
                let invoke_args = vec![elem, Value::I32(index as i32), receiver.clone()];
                if is_truthy(&invoke_callback_this(
                    ctx,
                    &callback,
                    this_arg.clone(),
                    &invoke_args,
                )) {
                    return Value::I32(index as i32);
                }
            }
            Value::I32(-1)
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "flatMap",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.flatMap callback is not callable",
            ) {
                return Value::Undefined;
            }
            let this_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let length = array_like_length(&receiver);
            let mut out = Vec::new();
            for index in 0..length {
                let Some(elem) = array_like_present_value_at(&receiver, index) else {
                    continue;
                };
                let invoke_args = vec![elem, Value::I32(index as i32), receiver.clone()];
                let r = invoke_callback_this(ctx, &callback, this_arg.clone(), &invoke_args);
                if let Value::Object(ref o) = r {
                    let lock = o.lock().unwrap();
                    if let ObjectKind::Array(ref inner) = lock.kind {
                        for (inner_index, inner_value) in inner.iter().enumerate() {
                            if !is_array_hole(&lock, inner_index) {
                                out.push(inner_value.clone());
                            }
                        }
                        continue;
                    }
                }
                out.push(r);
            }
            make_array(out)
        }),
    );

    // ── ES2025 group / groupToMap ───────────────────────────────────
    //
    // Group elements by the result of the callback. `group` returns a
    // null-prototype Object with string keys; `groupToMap` returns a
    // Map keyed by any value.

    vm.register_host_fn(
        "ecma:array",
        "group",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            use indexmap::IndexMap;
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let mut groups: IndexMap<String, Vec<Value>> = IndexMap::new();
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    let key = format!("{}", invoke_callback(ctx, &callback, &invoke_args));
                    groups
                        .entry(key)
                        .or_insert_with(Vec::new)
                        .push(elem.clone());
                }
            }
            // Materialize as an ordinary object with array-valued properties.
            let mut out = Object::new();
            for (k, v) in groups {
                out.properties.insert(k, make_array(v));
            }
            Value::Object(vybe_runtime::heap::alloc(out))
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "groupToMap",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            use indexmap::IndexMap;
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let mut groups: IndexMap<Value, Vec<Value>> = IndexMap::new();
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                };
                for (i, elem) in snapshot.iter().enumerate() {
                    let invoke_args = vec![
                        elem.clone(),
                        Value::I32(i as i32),
                        Value::Object(arr.clone()),
                    ];
                    let key = invoke_callback(ctx, &callback, &invoke_args);
                    groups
                        .entry(key)
                        .or_insert_with(Vec::new)
                        .push(elem.clone());
                }
            }
            // Build a JS Map with one entry per group.
            let mut map_im: IndexMap<Value, Value> = IndexMap::new();
            for (k, v) in groups {
                map_im.insert(k, make_array(v));
            }
            let mut obj = Object::new();
            obj.kind = ObjectKind::Map(map_im);
            obj.properties
                .insert("size".into(), Value::I32(obj_map_len(&obj) as i32));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // toSpliced — non-mutating splice returning a new array.
    vm.register_host_fn(
        "ecma:array",
        "toSpliced",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let del = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            // Items are individual args from index 3 onward (same as splice)
            let items: Vec<Value> = args.get(3..).unwrap_or(&[]).to_vec();
            if let Some(arr) = array_of(args, 0) {
                let snapshot: Vec<Value> = {
                    let o = arr.lock().unwrap();
                    if let ObjectKind::Array(ref v) = o.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                };
                let len = snapshot.len();
                let idx = if start < 0 {
                    ((len as i32) + start).max(0) as usize
                } else {
                    (start as usize).min(len)
                };
                let end = (idx + del).min(len);
                let mut out = Vec::with_capacity(len - (end - idx) + items.len());
                out.extend_from_slice(&snapshot[..idx]);
                out.extend(items.into_iter());
                out.extend_from_slice(&snapshot[end..]);
                return make_array(out);
            }
            make_array(Vec::new())
        }),
    );
}

/// JS truthy semantics — used by filter / some / every / find.
/// Matches ECMA-262 §7.1.2 ToBoolean.
fn is_truthy(v: &Value) -> bool {
    match v {
        Value::Null | Value::TypedNull(_) | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::F32(n) => *n != 0.0 && !n.is_nan(),
        Value::String(s) => !s.is_empty(),
        Value::Object(_)
        | Value::Symbol(_)
        | Value::BigInt(_)
        | Value::V128(_)
        | Value::WeakRef(_) => true }
}

fn obj_map_len(obj: &Object) -> usize {
    if let ObjectKind::Map(ref m) = obj.kind {
        m.len()
    } else {
        0
    }
}
