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

use crate::receiver_host_fn_ref;
use crate::typedarray::{
    fill_typed_array_bytes, new_typed_array, read_element, read_element_from_locked_buffer,
    reverse_typed_array_bytes, ta_live_length, typed_array_values_snapshot,
    write_array_values_to_typed_array_bytes, write_element,
};
use std::borrow::Cow;
use std::collections::{BTreeSet, HashSet};
use std::fmt::Write as _;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, TypedElemKind, Value};
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType};

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

fn invoke_callback(ctx: &mut HostContext, callback: &Value, args: &[Value]) -> Value {
    if let Some(v) = crate::function::invoke_bound_callback_if_needed(ctx, callback, args) {
        return v;
    }
    if let Some(v) = invoke_magic_callback(callback, args) {
        return v;
    }
    ctx.invoke(callback, args)
}

enum CallbackDispatch {
    Bound(crate::function::PreparedBoundCallback),
    Magic,
    Normal,
}

fn classify_callback(callback: &Value) -> CallbackDispatch {
    if let Some(bound) = crate::function::prepare_bound_callback(callback) {
        return CallbackDispatch::Bound(bound);
    }
    if let Value::Object(obj) = callback {
        if is_magic_callback_object(obj) {
            return CallbackDispatch::Magic;
        }
    }
    CallbackDispatch::Normal
}

fn invoke_classified_callback(
    ctx: &mut HostContext,
    callback: &Value,
    dispatch: &CallbackDispatch,
    args: &[Value],
) -> Value {
    match dispatch {
        CallbackDispatch::Bound(bound) => {
            crate::function::invoke_prepared_bound_callback(ctx, bound, args)
        }
        CallbackDispatch::Magic => {
            invoke_magic_callback(callback, args).unwrap_or(Value::Undefined)
        }
        CallbackDispatch::Normal => ctx.invoke(callback, args),
    }
}

fn is_callable_value(value: &Value) -> bool {
    match value {
        Value::Object(obj) => {
            matches!(
                obj.lock().unwrap().kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) || is_magic_callback_object(obj)
        }
        _ => false,
    }
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

fn invoke_classified_callback_this_ref(
    ctx: &mut HostContext,
    callback: &Value,
    dispatch: &CallbackDispatch,
    this_arg: &Value,
    args: &[Value],
) -> Value {
    match dispatch {
        CallbackDispatch::Bound(bound) => {
            crate::function::invoke_prepared_bound_callback(ctx, bound, args)
        }
        CallbackDispatch::Magic => {
            invoke_magic_callback(callback, args).unwrap_or(Value::Undefined)
        }
        CallbackDispatch::Normal => {
            crate::function::invoke_with_explicit_this(ctx, callback, this_arg.clone(), args)
        }
    }
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
        let acc = crate::keys::value_display_cow(&x);
        let cur = crate::keys::value_display_cow(&idx);
        let mut out = String::with_capacity(acc.len() + cur.len());
        out.push_str(acc.as_ref());
        out.push_str(cur.as_ref());
        return Some(owned_string_value(out));
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
        return Some(make_pair_array(x.clone(), x));
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
        _ => None,
    }
}

#[inline]
fn object_arg(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    match args.get(idx) {
        Some(Value::Object(obj)) => Some(obj.clone()),
        _ => None,
    }
}

pub(crate) fn make_array(elements: Vec<Value>) -> Value {
    let len = elements.len();
    let mut properties = vybe_runtime::value::Properties::with_capacity(2);
    properties.insert("length".into(), Value::F64(len as f64));
    properties.insert("__proto__".into(), shared_array_prototype());
    let obj = Object {
        properties,
        kind: ObjectKind::Array(elements),
        type_id: 0,
        fields: Vec::new(),
    };
    Value::Object(vybe_runtime::heap::alloc(obj))
}

#[inline]
fn make_array_from_args(args: &[Value]) -> Value {
    match args.len() {
        0 => make_array(Vec::new()),
        1 => make_array(vec![args[0].clone()]),
        _ => make_array(args.to_vec()),
    }
}

#[inline]
pub fn make_pair_array(first: Value, second: Value) -> Value {
    let mut properties = vybe_runtime::value::Properties::with_capacity(1);
    properties.insert("length".into(), Value::F64(2.0));
    Value::Object(vybe_runtime::heap::alloc(Object {
        properties,
        kind: ObjectKind::Array(vec![first, second]),
        type_id: 0,
        fields: Vec::new(),
    }))
}

#[inline]
fn array_entry_index_value(index: usize) -> Value {
    if index <= i32::MAX as usize {
        Value::I32(index as i32)
    } else {
        Value::F64(index as f64)
    }
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
        Some(Value::String(text)) => crate::keys::non_negative_integer_index_key(text),
        _ => None,
    }
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
            _ => None,
        })
        .collect()
}

fn hole_indices_opt(object: &Object) -> Option<BTreeSet<usize>> {
    let Some(Value::Object(holes)) = object.properties.get("__holes") else {
        return None;
    };
    let holes_guard = holes.lock().unwrap();
    let ObjectKind::Array(ref elems) = holes_guard.kind else {
        return None;
    };
    let mut out = BTreeSet::new();
    for value in elems {
        match value {
            Value::I32(index) if *index >= 0 => {
                out.insert(*index as usize);
            }
            Value::I64(index) if *index >= 0 => {
                out.insert(*index as usize);
            }
            _ => {}
        }
    }
    if out.is_empty() { None } else { Some(out) }
}

#[inline]
fn cached_hole_contains(holes: &Option<BTreeSet<usize>>, index: usize) -> bool {
    match holes {
        Some(holes) => holes.contains(&index),
        None => false,
    }
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
    let Some(Value::Object(holes)) = object.properties.get("__holes") else {
        return false;
    };
    let holes_guard = holes.lock().unwrap();
    let ObjectKind::Array(ref elems) = holes_guard.kind else {
        return false;
    };
    elems.iter().any(|value| match value {
        Value::I32(hole) if *hole >= 0 => *hole as usize == index,
        Value::I64(hole) if *hole >= 0 => *hole as usize == index,
        _ => false,
    })
}

pub fn mark_array_hole(object: &mut Object, index: usize) {
    if let Some(Value::Object(existing)) = object.properties.get("__holes") {
        let mut holes_guard = existing.lock().unwrap();
        if let ObjectKind::Array(ref mut elems) = holes_guard.kind {
            let exists = elems.iter().any(|value| match value {
                Value::I32(hole) if *hole >= 0 => *hole as usize == index,
                Value::I64(hole) if *hole >= 0 => *hole as usize == index,
                _ => false,
            });
            if !exists {
                elems.push(Value::I32(index as i32));
            }
            return;
        }
    }
    object.properties.insert(
        "__holes".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::I32(index as i32),
        ]))),
    );
}

pub fn clear_array_hole(object: &mut Object, index: usize) {
    let Some(Value::Object(existing)) = object.properties.get("__holes").cloned() else {
        return;
    };
    let mut empty = false;
    {
        let mut holes_guard = existing.lock().unwrap();
        if let ObjectKind::Array(ref mut elems) = holes_guard.kind {
            elems.retain(|value| match value {
                Value::I32(hole) if *hole >= 0 => *hole as usize != index,
                Value::I64(hole) if *hole >= 0 => *hole as usize != index,
                _ => true,
            });
            empty = elems.is_empty();
        }
    }
    if empty {
        object.properties.shift_remove("__holes");
    }
}

pub fn mark_hole_range(object: &mut Object, range: std::ops::Range<usize>) {
    if range.is_empty() {
        return;
    }
    if !object.properties.contains_key("__holes") {
        object.properties.insert(
            "__holes".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                range.map(|index| Value::I32(index as i32)).collect(),
            ))),
        );
        return;
    }
    let mut holes = hole_indices(object);
    holes.extend(range);
    store_hole_indices(object, &holes);
}

pub fn remap_array_holes<F>(object: &mut Object, mut remap: F)
where
    F: FnMut(usize) -> Option<usize>,
{
    let holes = hole_indices(object);
    let mut remapped = BTreeSet::new();
    for index in holes {
        if let Some(mapped) = remap(index) {
            remapped.insert(mapped);
        }
    }
    store_hole_indices(object, &remapped);
}

pub fn present_array_entries(object: &Object) -> Vec<(usize, Value)> {
    let ObjectKind::Array(ref values) = object.kind else {
        return Vec::new();
    };
    let holes = hole_indices_opt(object);
    if holes.is_none() {
        let mut entries = Vec::with_capacity(values.len());
        for (index, value) in values.iter().enumerate() {
            entries.push((index, value.clone()));
        }
        return entries;
    }
    let mut entries = Vec::with_capacity(values.len());
    for (index, value) in values.iter().enumerate() {
        if !cached_hole_contains(&holes, index) {
            entries.push((index, value.clone()));
        }
    }
    entries
}

fn array_length(arr: &Arc<Mutex<Object>>) -> usize {
    let o = arr.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(values) => values.len(),
        _ => 0,
    }
}

fn array_value_at(arr: &Arc<Mutex<Object>>, index: usize) -> Value {
    let o = arr.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(values) => values.get(index).cloned().unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

fn array_present_value_at(arr: &Arc<Mutex<Object>>, index: usize) -> Option<Value> {
    let o = arr.lock().unwrap();
    if is_array_hole(&o, index) {
        return None;
    }
    match &o.kind {
        ObjectKind::Array(values) => values.get(index).cloned(),
        _ => None,
    }
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
        _ => 0,
    }
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
        ObjectKind::Ordinary => crate::keys::with_index_key(index, |key| {
            object
                .properties
                .get(key)
                .cloned()
                .unwrap_or(Value::Undefined)
        }),
        ObjectKind::Map(map) => map
            .get_index(index)
            .map(|(_, value)| value.clone())
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
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
        ObjectKind::Ordinary => {
            crate::keys::with_index_key(index, |key| object.properties.get(key).cloned())
        }
        ObjectKind::Map(map) => map.get_index(index).map(|(_, value)| value.clone()),
        _ => None,
    }
}

fn array_like_dense_values(value: &Value) -> Vec<Value> {
    let Value::Object(obj) = value else {
        return Vec::new();
    };
    let object = obj.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(values) => values.clone(),
        ObjectKind::TypedArray(ta) => {
            let len = ta_live_length(ta);
            typed_array_values_snapshot(ta, len)
        }
        ObjectKind::Ordinary => {
            let length = property_length_as_usize(&object).unwrap_or(0);
            let mut values = Vec::with_capacity(length);
            for index in 0..length {
                values.push(crate::keys::with_index_key(index, |key| {
                    object
                        .properties
                        .get(key)
                        .cloned()
                        .unwrap_or(Value::Undefined)
                }));
            }
            values
        }
        ObjectKind::Map(map) => {
            let mut values = Vec::with_capacity(map.len());
            values.extend(map.values().cloned());
            values
        }
        _ => Vec::new(),
    }
}

fn typed_array_elem(value: &Value) -> Option<TypedElemKind> {
    let Value::Object(obj) = value else {
        return None;
    };
    let object = obj.lock().unwrap();
    match &object.kind {
        ObjectKind::TypedArray(ta) => Some(ta.elem),
        _ => None,
    }
}

fn make_typed_array_from_values(elem: TypedElemKind, values: &[Value]) -> Value {
    let typed = new_typed_array(elem, values.len());
    if let Value::Object(out) = &typed {
        let guard = out.lock().unwrap();
        if let ObjectKind::TypedArray(ta) = &guard.kind {
            let _ = write_array_values_to_typed_array_bytes(ta, 0, values);
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
        Some(crate::keys::value_display_string(value))
    }
}

fn compare_sort_values(
    ctx: &mut HostContext,
    compare_fn: Option<&Value>,
    compare_dispatch: Option<&CallbackDispatch>,
    a: &Value,
    b: &Value,
) -> std::cmp::Ordering {
    if let Some(compare_fn) = compare_fn.filter(|v| !matches!(v, Value::Undefined)) {
        let callback_args = [a.clone(), b.clone()];
        let result = if let Some(dispatch) = compare_dispatch {
            invoke_classified_callback(ctx, compare_fn, dispatch, &callback_args)
        } else {
            invoke_callback(ctx, compare_fn, &callback_args)
        };
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
            _ => std::cmp::Ordering::Equal,
        }
    }
}

fn sort_array_values_ecma(
    ctx: &mut HostContext,
    values: Vec<Value>,
    compare_fn: Option<&Value>,
) -> Vec<Value> {
    let len = values.len();
    let mut sortable = Vec::with_capacity(len);
    let mut undefined_count = 0usize;
    for value in values {
        if matches!(value, Value::Undefined) {
            undefined_count += 1;
        } else {
            sortable.push(value);
        }
    }
    if compare_fn
        .filter(|value| !matches!(value, Value::Undefined))
        .is_some()
    {
        let compare_dispatch = compare_fn
            .filter(|value| !matches!(value, Value::Undefined))
            .map(classify_callback);
        sortable
            .sort_by(|a, b| compare_sort_values(ctx, compare_fn, compare_dispatch.as_ref(), a, b));
    } else {
        let mut keyed: Vec<(String, Value)> = Vec::with_capacity(sortable.len());
        let mut can_use_keys = true;
        for value in sortable {
            if let Some(key) = default_sort_key(ctx, &value) {
                keyed.push((key, value));
            } else {
                can_use_keys = false;
                keyed.push((String::new(), value));
            }
        }
        if can_use_keys {
            keyed.sort_by(|a, b| a.0.cmp(&b.0));
            sortable = Vec::with_capacity(keyed.len());
            for (_, value) in keyed {
                sortable.push(value);
            }
        } else {
            sortable = Vec::with_capacity(keyed.len());
            for (_, value) in keyed {
                sortable.push(value);
            }
            sortable.sort_by(|a, b| compare_sort_values(ctx, None, None, a, b));
        }
    }
    sortable.extend((0..undefined_count).map(|_| Value::Undefined));
    sortable.resize(len, Value::Undefined);
    sortable
}

pub fn make_holey_array(length: usize) -> Value {
    let mut obj = Object::new_array(vec![Value::Undefined; length]);
    obj.properties
        .insert("__proto__".into(), shared_array_prototype());
    if length > 0 {
        let mut holes = Vec::with_capacity(length);
        for index in 0..length {
            holes.push(Value::I32(index as i32));
        }
        obj.properties.insert(
            "__holes".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(holes))),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(obj))
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
        Value::String(text) => crate::keys::non_negative_u32_key(text)
            .map(|length| length as usize)
            .ok_or("Invalid array length"),
        _ => Err("Invalid array length"),
    }
}

pub fn set_array_length(object: &mut Object, new_len: usize) {
    let old_len = match &object.kind {
        ObjectKind::Array(values) => values.len(),
        _ => return,
    };

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
        Err(message) => ctx.throw_value(crate::error::new_error(ctx, "RangeError", message)),
    }
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

/// Declare an `ecma:array` function — same closure, plus the signature.
///
/// No resource binding: a JS array is an ordinary object reference, not a
/// handle the host mints and drops, so `own`/`borrow` would claim a lifetime
/// that does not exist here.
fn array_fn(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new(
            "ecma:array",
            name,
            crate::perf::wrap("ecma:array", name, call),
        )
        .with_sig(FuncSig {
            name: name.to_string(),
            params: Param::unnamed_list(params),
            results,
        }),
    );
}

/// The array operand. §23.1.3 methods are generic over array-likes and several
/// handlers below accept a plain object carrying `length`, so the honest type
/// is `Any` — `list<T>` is homogeneous and copied, which an array reference is
/// not.
fn arr() -> ValType {
    ValType::Any
}

/// An arbitrary element. ECMA arrays are heterogeneous by definition.
fn elem() -> ValType {
    ValType::Any
}

/// A callback operand. The Component Model has no function type; a callable
/// marshals as an opaque value here.
fn callback() -> ValType {
    ValType::Any
}

/// An index. Negative values are meaningful (`at(-1)`), hence signed.
fn index() -> ValType {
    ValType::I32
}

/// An Array Iterator (§23.1.5) or an iterator result object. Both are ordinary
/// objects, so they type as `Any` for the same reason `arr()` does.
fn iterator() -> ValType {
    ValType::Any
}

/// WHY MOST OF §23.1.3 STAYS UNDECLARED.
///
/// A Component Model signature is a fixed parameter list: `declared_host_arity`
/// is `sig.params.len()`, and there is no spelling for an optional or a rest
/// parameter (`option<T>` is still positional). ECMA-262 array methods are
/// built on both — `slice([start[, end]])`, `filter(cb[, thisArg])`,
/// `push(...items)` — and every one of those call shapes is CORRECT, so a fixed
/// declaration would report correct callers as mismatches.
///
/// So the rule applied here is: declare a function only when the handler reads
/// no argument beyond its required ones. `None` keeps meaning UNKNOWN, never
/// zero, so leaving the rest undeclared costs nothing and claims nothing.
/// Widening `HostFnDecl` to express optionality is a runtime change and is not
/// in scope.
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
    fn_obj.properties.reserve(6);
    fn_obj.kind = ObjectKind::HostFunction(idx);
    fn_obj.properties.insert(
        "__host_module".into(),
        crate::keys::string_value("ecma:array"),
    );
    fn_obj
        .properties
        .insert("__host_name".into(), crate::keys::string_value(name));
    fn_obj
        .properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    fn_obj.properties.insert(
        "__proto__".into(),
        crate::function::shared_function_prototype(),
    );
    fn_obj
        .properties
        .insert("name".into(), crate::keys::string_value(name));
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
    let mut methods = Vec::with_capacity(METHODS.len());
    for name in METHODS {
        if let Some(method) = array_proto_method(vm, name) {
            methods.push((*name, method));
        }
    }
    {
        let mut p = proto.lock().unwrap();
        for (name, method) in methods {
            p.properties.insert(name.to_string(), method);
        }
        let nonenum = match p.properties.get("__nonenum") {
            Some(Value::Object(arr)) => arr.clone(),
            _ => {
                let arr = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
                p.properties
                    .insert("__nonenum".into(), Value::Object(arr.clone()));
                arr
            }
        };
        let mut nonenum = nonenum.lock().unwrap();
        if let ObjectKind::Array(ref mut elems) = nonenum.kind {
            for name in METHODS {
                if !elems
                    .iter()
                    .any(|value| matches!(value, Value::String(text) if text.as_ref() == *name))
                {
                    elems.push(crate::keys::string_value(name));
                }
            }
        }
    }
    // §17: a built-in data property is { [[Writable]]: true, [[Enumerable]]:
    // false, [[Configurable]]: true }. This list overlaps but is NOT identical
    // to the one `ecma_globals::populate` mounts, and only that one used to
    // mark anything — so precisely the four names unique to THIS list
    // (`keys`, `toLocaleString`, `toString`, `with`) came out enumerable and
    // fell out of `for (k in [])`, which the standard requires to yield the
    // indices alone. Marking at both mounts is what keeps the two lists from
    // disagreeing about attributes.
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
    array_fn(
        vm,
        "clear",
        vec![arr()],
        vec![],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_none() {
                    if let ObjectKind::Array(ref mut v) = o.kind {
                        v.clear();
                    }
                }
            }
            Value::Undefined
        }),
    );

    // first(arr) — `arr.at(0)`. Convenience for Queue.Peek.
    array_fn(
        vm,
        "first",
        vec![arr()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return v.first().cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );

    // last(arr) — `arr.at(-1)`. Convenience for Stack.Peek.
    array_fn(
        vm,
        "last",
        vec![arr()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return v.last().cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }),
    );

    // removeAt(arr, idx) — `arr.splice(idx, 1)`, returns removed value.
    array_fn(
        vm,
        "removeAt",
        vec![arr(), index()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_none() {
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
    array_fn(
        vm,
        "insertAt",
        vec![arr(), index(), elem()],
        vec![],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let idx = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_none() {
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
    array_fn(
        vm,
        "removeValue",
        vec![arr(), elem()],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_none() {
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
    vm.register_free_fn(
        "ecma:array",
        "new",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ `user_args`, NOT the raw list. `Array(4)` is a PLAIN call, and
            // under `ReceiverAbi::Parameter` every call puts a receiver at
            // argument 0 (§10.2.1.1 binds `undefined` for a plain one). Counted
            // as an element it made `Array(4)` a TWO-element array
            // `[undefined, 4]`, so `Array(4).fill(7)` produced `7,7`. Inert
            // under the ambient binding.
            let args = ctx.user_args(args, 0);
            match args.len() {
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
                    other => make_array(vec![other.clone()]),
                },
                _ => make_array_from_args(args),
            }
        }),
    );

    // newWithLength(n: i32) -> Array (n-element, null-filled).
    // Used by language-specific allocations (VB `ReDim`, .NET `new T[n]`)
    // that expect default-value semantics (null/0). JS callers go through
    // `new` above which materializes `undefined` slots per spec.
    array_fn(
        vm,
        "newWithLength",
        vec![index()],
        vec![arr()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args
                .first()
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            make_array(vec![Value::Null; n])
        }),
    );
    // Same handler under the Vybe-specific namespace — the emitters that
    // allocate a default-filled array (VB `ReDim`, .NET `new T[n]`) import it
    // from here, not from `ecma:array`.
    vm.register_host(
        HostFnDecl::new(
            "vybe:js-array",
            "newWithLength",
            Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                let n = args
                    .first()
                    .map(|v| v.as_i32().max(0) as usize)
                    .unwrap_or(0);
                make_array(vec![Value::Null; n])
            }),
        )
        .with_sig(FuncSig {
            name: "newWithLength".to_string(),
            params: Param::unnamed_list(vec![index()]),
            results: vec![arr()],
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| make_array_from_args(args)),
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
                            out.reserve(elems.len());
                            out.extend(elems.iter().cloned());
                        }
                        ObjectKind::TypedArray(ref ta) => {
                            let live = crate::typedarray::ta_live_length(ta);
                            out.reserve(live);
                            for index in 0..live {
                                out.push(read_element(ta, index));
                            }
                        }
                        // Map → Array of `[key, value]` pairs (§23.1.2.1).
                        ObjectKind::Map(ref m) => {
                            out.reserve(m.len());
                            for (k, v) in m.iter() {
                                out.push(crate::array::make_pair_array(k.clone(), v.clone()));
                            }
                        }
                        // Set → Array of values (§23.1.2.1).
                        ObjectKind::Set(ref set) => {
                            out.reserve(set.len());
                            out.extend(set.iter().cloned());
                        }
                        _ => {
                            if let Some(len_val) = s.properties.get("length") {
                                let len = len_val.as_f64().max(0.0) as usize;
                                out.reserve(len);
                                for i in 0..len {
                                    out.push(crate::keys::with_index_key(i, |key| {
                                        s.properties.get(key).cloned().unwrap_or(Value::Undefined)
                                    }));
                                }
                            }
                        }
                    }
                }
                Some(Value::String(s)) => {
                    out.reserve(s.len());
                    if s.is_ascii() {
                        for &byte in s.as_bytes() {
                            out.push(char_value(byte as char));
                        }
                    } else {
                        for c in s.chars() {
                            out.push(char_value(c));
                        }
                    }
                }
                _ => {}
            }
            if let Some(mapper) = args.get(1) {
                if !matches!(mapper, Value::Null | Value::Undefined) {
                    let mapper_dispatch = classify_callback(mapper);
                    let mut mapper_args = [Value::Undefined, Value::I32(0)];
                    for (i, slot) in out.iter_mut().enumerate() {
                        let v = std::mem::replace(slot, Value::Undefined);
                        mapper_args[0] = v;
                        mapper_args[1] = Value::I32(i as i32);
                        *slot =
                            invoke_classified_callback(ctx, mapper, &mapper_dispatch, &mapper_args);
                    }
                }
            }
            make_array(out)
        }),
    );

    // fromWithMap(arrayLike, mapFn) — Array.from with mandatory mapper.
    array_fn(
        vm,
        "fromWithMap",
        vec![elem(), callback()],
        vec![arr()],
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                match &s.kind {
                    ObjectKind::Array(elems) => {
                        out.reserve(elems.len());
                        out.extend(elems.iter().cloned());
                    }
                    _ => {
                        let len = s
                            .properties
                            .get("length")
                            .map(|v| v.as_f64().max(0.0) as usize)
                            .unwrap_or(0);
                        out.reserve(len);
                        for i in 0..len {
                            out.push(crate::keys::with_index_key(i, |key| {
                                s.properties.get(key).cloned().unwrap_or(Value::Undefined)
                            }));
                        }
                    }
                }
            }
            if let Some(mapper) = args.get(1) {
                let mapper_dispatch = classify_callback(mapper);
                let mut mapper_args = [Value::Undefined, Value::I32(0)];
                for (i, slot) in out.iter_mut().enumerate() {
                    let v = std::mem::replace(slot, Value::Undefined);
                    mapper_args[0] = v;
                    mapper_args[1] = Value::I32(i as i32);
                    *slot = invoke_classified_callback(ctx, mapper, &mapper_dispatch, &mapper_args);
                }
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
                Err(reason) => return crate::promise::make_promise("rejected", reason),
            };
            let mapper = args.get(1).cloned();
            let collected = crate::iterator::try_materialize_iterable_values(ctx, &source, true)
                .and_then(|values| {
                    let mut mapped = Vec::with_capacity(values.len());
                    let mut mapper_args = [Value::Undefined, Value::Undefined];
                    for (index, value) in values.into_iter().enumerate() {
                        let awaited = crate::iterator::try_maybe_await_value(value)?;
                        let mapped_value = match mapper.as_ref() {
                            Some(mapper) if !matches!(mapper, Value::Null | Value::Undefined) => {
                                mapper_args[0] = awaited;
                                mapper_args[1] = Value::I32(index as i32);
                                ctx.try_invoke(mapper, &mapper_args)?
                            }
                            _ => awaited,
                        };
                        mapped.push(crate::iterator::try_maybe_await_value(mapped_value)?);
                    }
                    Ok(make_array(mapped))
                });
            match collected {
                Ok(array) => crate::promise::make_promise("fulfilled", array),
                Err(reason) => crate::promise::make_promise("rejected", reason),
            }
        }),
    );

    // isArray(v) -> bool
    array_fn(
        vm,
        "isArray",
        vec![elem()],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            // §7.2.2 IsArray: a proxy is an Array when its target is.
            let mut v = args.first().cloned().unwrap_or(Value::Undefined);
            while let Value::Object(obj) = &v {
                match crate::object::proxy_target_and_handler(obj) {
                    Some((target, _)) => v = target,
                    None => break,
                }
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
    array_fn(
        vm,
        "get",
        vec![arr(), index()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let key = args.get(1).unwrap_or(&undefined);
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        if let Some(i) = crate::keys::non_negative_integer_index(key) {
                            return v.get(i).cloned().unwrap_or(Value::Undefined);
                        }
                        return Value::Undefined;
                    }
                    // Polymorphic dispatch on Map — the canonical cross-
                    // language associative type (PHP `['k'=>v]`, Python
                    // dicts, Ruby hashes, JS plain objects). Key is looked
                    // up using Value-level equality (SameValueZero), so
                    // both `$m['foo']` and `$m[$key]` where `$key = 'foo'`
                    // resolve identically.
                    ObjectKind::Map(m) => {
                        let lookup_key = crate::keys::map_lookup_key_cow(key);
                        if let Some(v) = m.get(lookup_key.as_ref()) {
                            return v.clone();
                        }
                        // PHP-ish fallback: if caller used a string key like
                        // "0" but the map stores integer keys (or vice
                        // versa), try the coerced form. Only coerces for
                        // purely numeric strings to avoid surprises.
                        if let Value::String(s) = key {
                            if let Some(n) = crate::keys::non_negative_i32_key(s) {
                                if let Some(v) = m.get(&Value::I32(n)) {
                                    return v.clone();
                                }
                            }
                        } else if let Value::I32(n) = key {
                            if *n >= 0 {
                                let string_key = crate::keys::small_index_string_value(*n as usize)
                                    .unwrap_or_else(|| {
                                        crate::keys::with_index_key(
                                            *n as usize,
                                            crate::keys::string_value,
                                        )
                                    });
                                if let Some(v) = m.get(&string_key).cloned() {
                                    return v;
                                }
                            } else {
                                let string_key = crate::keys::owned_string_value(n.to_string());
                                if let Some(v) = m.get(&string_key) {
                                    return v.clone();
                                }
                            }
                        }
                        return Value::Undefined;
                    }
                    _ => {}
                }
                // Plain Object fallback: property lookup. Used by Ordinary
                // objects and by compiler_common's `has` / `in` emitter.
                return crate::keys::with_property_key(key, |key_str| {
                    o.properties
                        .get(key_str)
                        .cloned()
                        .unwrap_or(Value::Undefined)
                });
            }
            Value::Undefined
        }),
    );

    // getValue(arr_or_obj, key) -> value. Same behavior as `get`, but the key
    // stays in the value ABI for VM collection op fallbacks that support string
    // and map keys in addition to numeric array indices.
    array_fn(
        vm,
        "getValue",
        vec![arr(), elem()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let key = args.get(1).unwrap_or(&undefined);
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        if let Some(i) = crate::keys::non_negative_integer_index(key) {
                            return v.get(i).cloned().unwrap_or(Value::Undefined);
                        }
                    }
                    ObjectKind::Map(m) => {
                        let lookup_key = crate::keys::map_lookup_key_cow(key);
                        if let Some(v) = m.get(lookup_key.as_ref()) {
                            return v.clone();
                        }
                        if let Value::String(s) = key {
                            if let Some(n) = crate::keys::non_negative_i32_key(s) {
                                if let Some(v) = m.get(&Value::I32(n)) {
                                    return v.clone();
                                }
                            }
                        } else if let Value::I32(n) = key {
                            if *n >= 0 {
                                let string_key = crate::keys::small_index_string_value(*n as usize)
                                    .unwrap_or_else(|| {
                                        crate::keys::with_index_key(
                                            *n as usize,
                                            crate::keys::string_value,
                                        )
                                    });
                                if let Some(v) = m.get(&string_key).cloned() {
                                    return v;
                                }
                            } else {
                                let string_key = crate::keys::owned_string_value(n.to_string());
                                if let Some(v) = m.get(&string_key) {
                                    return v.clone();
                                }
                            }
                        }
                        return Value::Undefined;
                    }
                    ObjectKind::TypedArray(ta) => {
                        if let Some(idx) = crate::keys::non_negative_integer_index(key) {
                            return read_element(ta, idx);
                        }
                    }
                    _ => {}
                }
                return crate::keys::with_property_key(key, |key_str| {
                    o.properties
                        .get(key_str)
                        .cloned()
                        .unwrap_or(Value::Undefined)
                });
            }
            Value::Undefined
        }),
    );

    // set(arr_or_obj, key, v) -> () — extends arrays with null-fill when
    // key >= length; stores into plain objects by string key; updates Maps
    // using the canonical Value-keyed IndexMap.
    array_fn(
        vm,
        "set",
        vec![arr(), index(), elem()],
        vec![],
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let key = args.get(1).unwrap_or(&undefined);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                if matches!(&o.kind, ObjectKind::Array(_))
                    && matches!(key, Value::String(text) if text.as_ref() == "length" || text.as_ref() == "__len__")
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
                        if let Some(idx) = crate::keys::non_negative_integer_index(key) {
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
                            crate::keys::with_property_key(key, |key_str| {
                                o.properties.insert(key_str.to_owned(), val);
                            });
                        }
                    }
                    ObjectKind::Map(m) => {
                        let map_key = crate::keys::map_lookup_key(key);
                        m.insert(map_key, val);
                    }
                    ObjectKind::TypedArray(ta) => {
                        if let Some(idx) = crate::keys::non_negative_integer_index(key) {
                            write_element(ta, idx, &val);
                        }
                    }
                    _ => {
                        crate::keys::with_property_key(key, |key_str| {
                            o.properties.insert(key_str.to_owned(), val);
                        });
                    }
                }
            }
            Value::Null
        }),
    );

    // setValue(arr_or_obj, key, v) -> (). Value-key companion to `set`.
    array_fn(
        vm,
        "setValue",
        vec![arr(), elem(), elem()],
        vec![],
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let key = args.get(1).unwrap_or(&undefined);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                if matches!(&o.kind, ObjectKind::Array(_))
                    && matches!(key, Value::String(text) if text.as_ref() == "length" || text.as_ref() == "__len__")
                {
                    apply_js_array_length(ctx, &mut o, &val);
                    return Value::Null;
                }
                match &mut o.kind {
                    ObjectKind::Array(v) => {
                        if let Some(idx) = crate::keys::non_negative_integer_index(key) {
                            let old_len = v.len();
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
                            crate::keys::with_property_key(key, |key_str| {
                                o.properties.insert(key_str.to_owned(), val);
                            });
                        }
                    }
                    ObjectKind::Map(m) => {
                        let map_key = crate::keys::map_lookup_key(key);
                        m.insert(map_key, val);
                    }
                    ObjectKind::TypedArray(ta) => {
                        if let Some(idx) = crate::keys::non_negative_integer_index(key) {
                            write_element(ta, idx, &val);
                        }
                    }
                    _ => {
                        crate::keys::with_property_key(key, |key_str| {
                            o.properties.insert(key_str.to_owned(), val);
                        });
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
    array_fn(
        vm,
        "length",
        vec![arr()],
        vec![index()],
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
                        .unwrap_or(Value::Null),
                };
            }
            Value::Null
        }),
    );

    // setLength(arr, n) -> () — truncate or null-fill extend
    array_fn(
        vm,
        "setLength",
        vec![arr(), index()],
        vec![],
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
    array_fn(
        vm,
        "at",
        vec![arr(), index()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let i = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if let Some(arr) = object_arg(args, 0) {
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
                    _ => 0,
                };
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
                        crate::keys::with_index_key(start + offset, |key| {
                            object.properties.insert(key.to_owned(), value.clone());
                        });
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
    array_fn(
        vm,
        "pop",
        vec![arr()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_some() {
                    return Value::Undefined;
                }
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
    array_fn(
        vm,
        "shift",
        vec![arr()],
        vec![elem()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_some() {
                    return Value::Undefined;
                }
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
                            other => Some(other - 1),
                        });
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
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if o.properties.get(FROZEN_MARK).is_some() {
                    if let ObjectKind::Array(ref v) = o.kind {
                        return Value::I32(v.len() as i32);
                    }
                    return Value::I32(0);
                }
                let offset = args.len().saturating_sub(1);
                let len = if let ObjectKind::Array(ref mut v) = o.kind {
                    for (i, val) in args.iter().skip(1).enumerate() {
                        v.insert(i, val.clone());
                    }
                    v.len() as i32
                } else {
                    return Value::I32(0);
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
                Some(v) => v.as_i32().max(0) as usize,
            };
            let items = args.get(3..).unwrap_or(&[]);
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
                    deleted.reserve(delete_count);
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
                        for (i, val) in items.iter().cloned().enumerate() {
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
    array_fn(
        vm,
        "reverse",
        vec![arr()],
        vec![arr()],
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
                        let mut values = typed_array_values_snapshot(ta, live);
                        values = sort_array_values_ecma(ctx, values, compare_fn.as_ref());
                        let _ = write_array_values_to_typed_array_bytes(ta, 0, &values);
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
            // §23.1.3.7 step 7: "If end is undefined, let relativeEnd be
            // len" — the same rule `slice` below already follows.
            let end = match args.get(3) {
                None | Some(Value::Undefined) => i32::MAX,
                Some(v) => v.as_i32(),
            };
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
                        if fill_typed_array_bytes(ta, s, e, &val) {
                            return args.first().cloned().unwrap_or(Value::Null);
                        }
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
            // §23.1.3.4 step 8: "If end is undefined, let relativeEnd be
            // len" — the same rule `slice` below already follows.
            let end = match args.get(3) {
                None | Some(Value::Undefined) => i32::MAX,
                Some(v) => v.as_i32(),
            };
            // §23.1.3.4 relative indexing — same normalization as fill.
            let norm = |x: i32, len: i32| -> usize {
                if x < 0 {
                    ((len as i64) + (x as i64)).max(0) as usize
                } else {
                    x.min(len) as usize
                }
            };
            if let Some(arr) = object_arg(args, 0) {
                let mut o = arr.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = o.kind {
                    let len = v.len() as i32;
                    let t = norm(target, len);
                    let s = norm(start, len);
                    let e = norm(end, len).max(s);
                    let mut slice = Vec::with_capacity(e.saturating_sub(s));
                    for value in &v[s..e] {
                        slice.push(value.clone());
                    }
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
            if let Some(arr) = object_arg(args, 0) {
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
            // §23.1.3.27: `undefined` is not 0 and not "no argument" — for
            // `start` it means 0 and for `end` it means `len`. Reading it
            // through `as_i32()` gave BOTH of them 0, so an explicitly-passed
            // `undefined` end truncated the slice to empty while an OMITTED
            // one returned the whole thing. Those must agree, because a WASM
            // import has one arity: a call site cannot choose to omit.
            let start = match args.get(1) {
                None | Some(Value::Undefined) => 0,
                Some(v) => v.as_i32(),
            };
            let end = match args.get(2) {
                None | Some(Value::Undefined) => i32::MAX,
                Some(v) => v.as_i32(),
            };
            if let Some(Value::String(s)) = args.first() {
                if s.is_ascii() {
                    let len = s.len() as i32;
                    let si = (if start < 0 { len + start } else { start })
                        .max(0)
                        .min(len) as usize;
                    let ei = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                    if si < ei {
                        return crate::keys::string_value(&s[si..ei]);
                    }
                    return crate::keys::string_value("");
                }
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
                return owned_string_value(out);
            }
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let s = (if start < 0 { len + start } else { start })
                        .max(0)
                        .min(len) as usize;
                    let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                    let out: Vec<Value> = if s < e {
                        let mut out = Vec::with_capacity(e - s);
                        for value in &v[s..e] {
                            out.push(value.clone());
                        }
                        out
                    } else {
                        Vec::new()
                    };
                    let sliced = make_array(out);
                    if let Value::Object(obj) = &sliced {
                        let mut holes = BTreeSet::new();
                        let old_holes = hole_indices_opt(&o);
                        for index in s..e {
                            if cached_hole_contains(&old_holes, index) {
                                holes.insert(index - s);
                            }
                        }
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
    // Spec `concat` is variadic; this handler reads exactly one other operand,
    // so the declaration states what it implements, not what §23.1.3.1 allows.
    array_fn(
        vm,
        "concat",
        vec![arr(), arr()],
        vec![arr()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut out = Vec::new();
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    out.reserve(v.len());
                    out.extend(v.iter().cloned());
                }
            }
            // §23.1.3.1 Array.prototype.concat step 5.b — IsConcatSpreadable.
            //
            // ⛔ THE QUESTION IS NOT "IS IT ITERABLE". §22.1.3.1.1 says: a
            // non-Object is NEVER spread; otherwise ask `@@isConcatSpreadable`,
            // and only when that is `undefined` fall back to `IsArray`. This
            // used to spread by KIND, which got three things wrong at once —
            // a STRING was spread into its code points (`[].concat("ab")` gave
            // `a,b` where the spec appends `"ab"` whole), an array-like opting
            // IN with the symbol was appended as one element, and an array
            // opting OUT with `false` was spread anyway.
            match args.get(1) {
                Some(Value::Object(o)) => {
                    let (explicit, is_array, length) = {
                        let lock = o.lock().unwrap();
                        let explicit = lock.properties.get("isconcatspreadable").cloned();
                        let is_array = matches!(lock.kind, ObjectKind::Array(_));
                        let len = lock
                            .properties
                            .get("length")
                            .map(|v| v.as_f64() as usize)
                            .unwrap_or(0);
                        (explicit, is_array, len)
                    };
                    let spreadable = match explicit {
                        Some(Value::Undefined) | None => is_array,
                        // ⛔ ToBoolean (§7.1.2), NOT `Value::as_bool`. That
                        // accessor is `matches!(self, Bool(true))` — it answers
                        // "is this the boolean true", so every truthy NON-bool
                        // read as false. §22.1.3.1.1 step 4 applies ToBoolean,
                        // so `arr[@@isConcatSpreadable] = 1` must spread just
                        // as `= true` does; measured, `1` did not and `true`
                        // did.
                        Some(v) => crate::boolean::to_boolean(&v),
                    };
                    if !spreadable {
                        out.push(Value::Object(o.clone()));
                    } else {
                        let lock = o.lock().unwrap();
                        match &lock.kind {
                            ObjectKind::Array(v) => {
                                out.reserve(v.len());
                                out.extend(v.iter().cloned());
                            }
                            ObjectKind::Set(sv) => {
                                out.reserve(sv.len());
                                out.extend(sv.iter().cloned());
                            }
                            ObjectKind::Map(m) => {
                                out.reserve(m.len());
                                for (k, v) in m {
                                    out.push(crate::array::make_pair_array(k.clone(), v.clone()));
                                }
                            }
                            _ => {
                                // Array-LIKE that opted in: spread by `length`, reading
                                // the indexed properties (§23.1.3.1 step 5.c.iv).
                                out.reserve(length);
                                for i in 0..length {
                                    out.push(crate::keys::with_index_key(i, |key| {
                                        lock.properties
                                            .get(key)
                                            .cloned()
                                            .unwrap_or(Value::Undefined)
                                    }));
                                }
                            }
                        }
                    }
                }
                // A primitive is never spread — including a String, whose
                // `Symbol.iterator` is irrelevant here.
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
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let start = from.max(0) as usize;
                    let holes = hole_indices_opt(&o);
                    for (i, elem) in v.iter().enumerate().skip(start) {
                        if cached_hole_contains(&holes, i) {
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
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let len = v.len() as i32;
                    let from = args.get(2).map(|v| v.as_i32()).unwrap_or(len - 1);
                    let end = from.min(len - 1).max(-1);
                    let end_idx = if end < 0 { 0 } else { (end + 1) as usize };
                    let holes = hole_indices_opt(&o);
                    for (i, elem) in v[..end_idx].iter().enumerate().rev() {
                        if cached_hole_contains(&holes, i) {
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
                        let holes = hole_indices_opt(&o);
                        for (index, elem) in v.iter().enumerate().skip(from) {
                            if cached_hole_contains(&holes, index) {
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
                .map(|v| match v {
                    Value::String(text) => Cow::Borrowed(text.as_ref()),
                    other => crate::keys::value_display_cow(other),
                })
                .unwrap_or(Cow::Borrowed(","));
            let mut joined = String::new();
            let mut first = true;
            if let Some(Value::Object(o)) = args.first() {
                let inner = o.lock().unwrap();
                match &inner.kind {
                    ObjectKind::Array(v) => {
                        let holes = hole_indices_opt(&inner);
                        for (index, value) in v.iter().enumerate() {
                            if cached_hole_contains(&holes, index) {
                                push_join_part(&mut joined, &sep, &mut first, "");
                            } else {
                                push_join_value(&mut joined, &sep, &mut first, value);
                            }
                        }
                    }
                    ObjectKind::Map(m) => {
                        for value in m.values() {
                            push_join_value(&mut joined, &sep, &mut first, value);
                        }
                    }
                    ObjectKind::Ordinary => {
                        if let Some(len) = property_length_as_usize(&inner) {
                            for index in 0..len {
                                if let Some(value) = crate::keys::with_index_key(index, |key| {
                                    inner.properties.get(key)
                                }) {
                                    push_join_value(&mut joined, &sep, &mut first, value);
                                } else {
                                    push_join_part(&mut joined, &sep, &mut first, "");
                                }
                            }
                        } else {
                            // Plain JS object — iterate values in insertion
                            // order. The compiler tracks insertion order in a
                            // side `__keys` array when index-assigning string
                            // keys; without it, `properties` is a HashMap and
                            // iteration order is randomized per process.
                            let mut used_ordered_keys = false;
                            if let Some(Value::Object(arr)) = inner.properties.get("__keys") {
                                let a = arr.lock().unwrap();
                                if let ObjectKind::Array(items) = &a.kind {
                                    used_ordered_keys = true;
                                    for key_value in items {
                                        crate::keys::with_property_key(key_value, |key| {
                                            if !key.starts_with("__") {
                                                if let Some(value) = inner.properties.get(key) {
                                                    push_join_value(
                                                        &mut joined,
                                                        &sep,
                                                        &mut first,
                                                        value,
                                                    );
                                                }
                                            }
                                        });
                                    }
                                }
                            }
                            if !used_ordered_keys {
                                for (_, value) in inner
                                    .properties
                                    .iter()
                                    .filter(|(k, _)| !k.starts_with("__"))
                                {
                                    push_join_value(&mut joined, &sep, &mut first, value);
                                }
                            }
                        }
                    }
                    ObjectKind::TypedArray(ta) => {
                        let live = ta_live_length(ta);
                        let bpe = ta.elem.bytes_per_element();
                        let buf = ta.buffer.lock().unwrap();
                        for i in 0..live {
                            let abs = ta.byte_offset + i * bpe;
                            push_join_sep(&mut joined, &sep, &mut first);
                            push_typed_array_join_element(
                                &mut joined,
                                read_element_from_locked_buffer(ta.elem, &buf, abs, bpe),
                            );
                        }
                    }
                    _ => {}
                }
            }
            owned_string_value(joined)
        }),
    );

    // toString(arr) -> string (same as join with default ",")
    array_fn(
        vm,
        "toString",
        vec![arr()],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let mut joined = String::new();
                    let mut first = true;
                    let holes = hole_indices_opt(&o);
                    for (index, value) in v.iter().enumerate() {
                        if cached_hole_contains(&holes, index) {
                            push_join_part(&mut joined, ",", &mut first, "");
                        } else {
                            push_join_value(&mut joined, ",", &mut first, value);
                        }
                    }
                    return owned_string_value(joined);
                }
            }
            crate::keys::string_value("")
        }),
    );

    // toLocaleString — same as toString for MVP
    array_fn(
        vm,
        "toLocaleString",
        vec![arr()],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            // Same as toString — real locale-aware conversion lives in
            // Phase F (intl integration).
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let mut joined = String::new();
                    let mut first = true;
                    let holes = hole_indices_opt(&o);
                    for (index, value) in v.iter().enumerate() {
                        if cached_hole_contains(&holes, index) {
                            push_join_part(&mut joined, ",", &mut first, "");
                        } else {
                            push_join_value(&mut joined, ",", &mut first, value);
                        }
                    }
                    return owned_string_value(joined);
                }
            }
            crate::keys::string_value("")
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
                let holes = hole_indices_opt(arr_obj);
                for (i, v) in elems.iter().enumerate() {
                    if cached_hole_contains(&holes, i) {
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
            if let Some(arr) = object_arg(args, 0) {
                let id = Arc::as_ptr(&arr) as usize;
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    out.reserve(v.len());
                    let mut seen = HashSet::from([id]);
                    flatten(ctx, &mut out, &o, v, depth, &mut seen);
                }
            }
            make_array(out)
        }),
    );

    // ── ES2023 non-mutating variants ────────────────────────────────

    // toReversed(arr) -> new_arr
    array_fn(
        vm,
        "toReversed",
        vec![arr()],
        vec![arr()],
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
    array_fn(
        vm,
        "with",
        vec![arr(), index(), elem()],
        vec![arr()],
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

fn push_typed_array_join_element(out: &mut String, value: Value) {
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

fn push_join_sep(out: &mut String, sep: &str, first: &mut bool) {
    if *first {
        *first = false;
    } else {
        out.push_str(sep);
    }
}

fn push_join_part(out: &mut String, sep: &str, first: &mut bool, part: &str) {
    push_join_sep(out, sep, first);
    out.push_str(part);
}

fn push_join_value(out: &mut String, sep: &str, first: &mut bool, value: &Value) {
    match value {
        Value::Null | Value::Undefined => push_join_part(out, sep, first, ""),
        Value::String(text) => push_join_part(out, sep, first, text.as_ref()),
        Value::BigInt(n) => {
            push_join_sep(out, sep, first);
            let _ = write!(out, "{}", n);
        }
        other => {
            push_join_sep(out, sep, first);
            let _ = write!(out, "{}", other);
        }
    }
}

// ── Iteration / higher-order callbacks ─────────────────────────────────

/// Captured at register-time so `make_array_iterator` can stamp a
/// HostFunction property pointing at `iterNext` without re-resolving
/// the registry on every call.
static ARRAY_ITER_NEXT_IDX: std::sync::OnceLock<usize> = std::sync::OnceLock::new();
static ARRAY_ITER_NEXT_METHOD: std::sync::OnceLock<Value> = std::sync::OnceLock::new();

#[inline]
fn array_iter_next_host_key() -> &'static (String, String) {
    static KEY: std::sync::OnceLock<(String, String)> = std::sync::OnceLock::new();
    KEY.get_or_init(|| ("ecma:array".to_string(), "iterNext".to_string()))
}

fn iter_result(value: Value, done: bool) -> Value {
    let mut obj = Object::new();
    obj.properties.reserve(2);
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
    obj.properties
        .reserve(if ARRAY_ITER_NEXT_IDX.get().is_some() {
            3
        } else {
            2
        });
    obj.kind = ObjectKind::Array(materialized);
    obj.properties
        .insert("__type".into(), crate::keys::string_value("ArrayIterator"));
    obj.properties.insert("__index".into(), Value::I32(0));
    if let Some(idx) = ARRAY_ITER_NEXT_IDX.get() {
        obj.properties.insert(
            "next".into(),
            ARRAY_ITER_NEXT_METHOD
                .get_or_init(|| receiver_host_fn_ref("ecma:array", "iterNext", *idx))
                .clone(),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn register_iteration(vm: &mut VM) {
    // `iterNext(this)` — implements §23.1.5.2.1 Array Iterator next().
    // Reads `__index`, returns `{value, done}`, advances the cursor.
    array_fn(
        vm,
        "iterNext",
        vec![iterator()],
        vec![iterator()],
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
    if let Some(idx) = vm.host_registry.get(array_iter_next_host_key()).copied() {
        let _ = ARRAY_ITER_NEXT_IDX.set(idx);
        let _ = ARRAY_ITER_NEXT_METHOD.set(receiver_host_fn_ref("ecma:array", "iterNext", idx));
    }

    // keys(arr) / values(arr) / entries(arr) — §23.1.3.{16,36,7}.
    // Return a §23.1.5 Array Iterator with `next()` driving the cursor.
    // The iterator's underlying Array kind keeps spread / for-of working
    // through plain-array iteration paths.
    array_fn(
        vm,
        "keys",
        vec![arr()],
        vec![iterator()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    let mut out = Vec::with_capacity(v.len());
                    for i in 0..v.len() {
                        out.push(array_entry_index_value(i));
                    }
                    return make_array_iterator(out);
                }
            }
            make_array_iterator(Vec::new())
        }),
    );

    array_fn(
        vm,
        "values",
        vec![arr()],
        vec![iterator()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(arr) = object_arg(args, 0) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    return make_array_iterator(v.clone());
                }
            }
            make_array_iterator(Vec::new())
        }),
    );

    array_fn(
        vm,
        "entries",
        vec![arr()],
        vec![iterator()],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let mut out = Vec::with_capacity(v.len());
                        for (i, e) in v.iter().enumerate() {
                            out.push(crate::array::make_pair_array(
                                array_entry_index_value(i),
                                e.clone(),
                            ));
                        }
                        return make_array_iterator(out);
                    }
                    ObjectKind::TypedArray(ta) => {
                        let live = ta_live_length(ta);
                        let values = typed_array_values_snapshot(ta, live);
                        let mut out = Vec::with_capacity(live);
                        for (i, value) in values.into_iter().enumerate() {
                            out.push(make_array(vec![array_entry_index_value(i), value]));
                        }
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

    array_fn(
        vm,
        "forEach",
        vec![arr(), callback()],
        vec![],
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let dispatch = classify_callback(&callback);
            if let Some(arr) = object_arg(args, 0) {
                let length = array_length(&arr);
                let receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver];
                for index in 0..length {
                    let Some(elem) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(index as i32);
                    invoke_classified_callback(ctx, &callback, &dispatch, &invoke_args);
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "map",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.map callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let mut mapped_values = vec![Value::Undefined; length];
                let mut holes = Vec::new();
                let array_receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), array_receiver];
                for index in 0..length {
                    let Some(elem) = array_present_value_at(&arr, index) else {
                        holes.push(Value::I32(index as i32));
                        continue;
                    };
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(index as i32);
                    mapped_values[index] = invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    );
                }
                let mut mapped = Object::new_array(mapped_values);
                mapped
                    .properties
                    .insert("__proto__".into(), shared_array_prototype());
                if !holes.is_empty() {
                    mapped.properties.insert(
                        "__holes".into(),
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(holes))),
                    );
                }
                return Value::Object(vybe_runtime::heap::alloc(mapped));
            }
            let length = array_like_length(&receiver);
            if length > 0 || matches!(receiver, Value::Object(_)) {
                let mut mapped = Vec::with_capacity(length);
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
                for i in 0..length {
                    let elem = array_like_value_at(&receiver, i);
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(i as i32);
                    mapped.push(invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    ));
                }
                return make_array(mapped);
            }
            make_array(Vec::new())
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "filter",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.filter callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let mut filtered = Vec::with_capacity(length);
                let array_receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), array_receiver];
                for index in 0..length {
                    let Some(elem) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    invoke_args[0] = elem.clone();
                    invoke_args[1] = Value::I32(index as i32);
                    if is_truthy(&invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    )) {
                        filtered.push(elem);
                    }
                }
                return make_array(filtered);
            }
            let length = array_like_length(&receiver);
            if length > 0 || matches!(receiver, Value::Object(_)) {
                let mut filtered = Vec::with_capacity(length);
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
                for i in 0..length {
                    let elem = array_like_value_at(&receiver, i);
                    invoke_args[0] = elem.clone();
                    invoke_args[1] = Value::I32(i as i32);
                    if is_truthy(&invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    )) {
                        filtered.push(elem);
                    }
                }
                return make_array(filtered);
            }
            make_array(Vec::new())
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "reduce",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
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
            let dispatch = classify_callback(&callback);
            if let Some(arr) = object_arg(args, 0) {
                let length = array_length(&arr);
                let receiver = Value::Object(arr.clone());
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
                let mut invoke_args = [Value::Undefined, Value::Undefined, Value::I32(0), receiver];
                for index in start_idx..length {
                    let Some(value) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    invoke_args[0] = acc;
                    invoke_args[1] = value;
                    invoke_args[2] = Value::I32(index as i32);
                    acc = invoke_classified_callback(ctx, &callback, &dispatch, &invoke_args);
                }
            }
            acc
        }),
    );

    vm.register_host_fn(
        "ecma:array",
        "reduceRight",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let initial_provided =
                args.len() > 2 && !matches!(args.get(2), Some(Value::Undefined) | None);
            let mut acc = if initial_provided {
                args.get(2).cloned().unwrap_or(Value::Undefined)
            } else {
                Value::Undefined
            };
            let dispatch = classify_callback(&callback);
            if let Some(arr) = object_arg(args, 0) {
                let length = array_length(&arr);
                let receiver = Value::Object(arr.clone());
                let mut end_idx = length;
                if !initial_provided {
                    let mut found = false;
                    for index in (0..length).rev() {
                        if let Some(value) = array_present_value_at(&arr, index) {
                            acc = value;
                            end_idx = index;
                            found = true;
                            break;
                        }
                    }
                    if !found {
                        return Value::Undefined;
                    }
                }
                let mut invoke_args = [Value::Undefined, Value::Undefined, Value::I32(0), receiver];
                for index in (0..end_idx).rev() {
                    let Some(value) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    invoke_args[0] = acc;
                    invoke_args[1] = value;
                    invoke_args[2] = Value::I32(index as i32);
                    acc = invoke_classified_callback(ctx, &callback, &dispatch, &invoke_args);
                }
            }
            acc
        }),
    );

    array_fn(
        vm,
        "some",
        vec![arr(), callback()],
        vec![ValType::Bool],
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let dispatch = classify_callback(&callback);
            if let Some(arr) = object_arg(args, 0) {
                let length = array_length(&arr);
                let receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver];
                for index in 0..length {
                    let Some(elem) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(index as i32);
                    if is_truthy(&invoke_classified_callback(
                        ctx,
                        &callback,
                        &dispatch,
                        &invoke_args,
                    )) {
                        return Value::Bool(true);
                    }
                }
            }
            Value::Bool(false)
        }),
    );

    array_fn(
        vm,
        "every",
        vec![arr(), callback()],
        vec![ValType::Bool],
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let dispatch = classify_callback(&callback);
            if let Some(arr) = object_arg(args, 0) {
                let length = array_length(&arr);
                let receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver];
                for index in 0..length {
                    let Some(elem) = array_present_value_at(&arr, index) else {
                        continue;
                    };
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(index as i32);
                    if !is_truthy(&invoke_classified_callback(
                        ctx,
                        &callback,
                        &dispatch,
                        &invoke_args,
                    )) {
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
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.find callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let array_receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), array_receiver];
                for index in 0..length {
                    let elem = array_value_at(&arr, index);
                    invoke_args[0] = elem.clone();
                    invoke_args[1] = Value::I32(index as i32);
                    if is_truthy(&invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    )) {
                        return elem.clone();
                    }
                }
                return Value::Undefined;
            }
            let length = array_like_length(&receiver);
            let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
            for index in 0..length {
                let elem = array_like_value_at(&receiver, index);
                invoke_args[0] = elem.clone();
                invoke_args[1] = Value::I32(index as i32);
                if is_truthy(&invoke_classified_callback_this_ref(
                    ctx,
                    &callback,
                    &dispatch,
                    this_arg,
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
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.findIndex callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let array_receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), array_receiver];
                for index in 0..length {
                    let elem = array_value_at(&arr, index);
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(index as i32);
                    if is_truthy(&invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    )) {
                        return Value::I32(index as i32);
                    }
                }
                return Value::I32(-1);
            }
            let length = array_like_length(&receiver);
            let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
            for index in 0..length {
                let elem = array_like_value_at(&receiver, index);
                invoke_args[0] = elem;
                invoke_args[1] = Value::I32(index as i32);
                if is_truthy(&invoke_classified_callback_this_ref(
                    ctx,
                    &callback,
                    &dispatch,
                    this_arg,
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
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.findLast callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let array_receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), array_receiver];
                for index in (0..length).rev() {
                    let elem = array_value_at(&arr, index);
                    invoke_args[0] = elem.clone();
                    invoke_args[1] = Value::I32(index as i32);
                    if is_truthy(&invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    )) {
                        return elem.clone();
                    }
                }
                return Value::Undefined;
            }
            let length = array_like_length(&receiver);
            let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
            for index in (0..length).rev() {
                let elem = array_like_value_at(&receiver, index);
                invoke_args[0] = elem.clone();
                invoke_args[1] = Value::I32(index as i32);
                if is_truthy(&invoke_classified_callback_this_ref(
                    ctx,
                    &callback,
                    &dispatch,
                    this_arg,
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
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.findLastIndex callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                let array_receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), array_receiver];
                for index in (0..length).rev() {
                    let elem = array_value_at(&arr, index);
                    invoke_args[0] = elem;
                    invoke_args[1] = Value::I32(index as i32);
                    if is_truthy(&invoke_classified_callback_this_ref(
                        ctx,
                        &callback,
                        &dispatch,
                        this_arg,
                        &invoke_args,
                    )) {
                        return Value::I32(index as i32);
                    }
                }
                return Value::I32(-1);
            }
            let length = array_like_length(&receiver);
            let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
            for index in (0..length).rev() {
                let elem = array_like_value_at(&receiver, index);
                invoke_args[0] = elem;
                invoke_args[1] = Value::I32(index as i32);
                if is_truthy(&invoke_classified_callback_this_ref(
                    ctx,
                    &callback,
                    &dispatch,
                    this_arg,
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
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            if !require_callable(
                ctx,
                &callback,
                "Array.prototype.flatMap callback is not callable",
            ) {
                return Value::Undefined;
            }
            let undefined = Value::Undefined;
            let this_arg = args.get(2).unwrap_or(&undefined);
            let dispatch = classify_callback(&callback);
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let length = array_like_length(&receiver);
            let mut out = Vec::with_capacity(length);
            let mut invoke_args = [Value::Undefined, Value::I32(0), receiver.clone()];
            for index in 0..length {
                let Some(elem) = array_like_present_value_at(&receiver, index) else {
                    continue;
                };
                invoke_args[0] = elem;
                invoke_args[1] = Value::I32(index as i32);
                let r = invoke_classified_callback_this_ref(
                    ctx,
                    &callback,
                    &dispatch,
                    this_arg,
                    &invoke_args,
                );
                if let Value::Object(ref o) = r {
                    let lock = o.lock().unwrap();
                    if let ObjectKind::Array(ref inner) = lock.kind {
                        let holes = hole_indices_opt(&lock);
                        for (inner_index, inner_value) in inner.iter().enumerate() {
                            if !cached_hole_contains(&holes, inner_index) {
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

    array_fn(
        vm,
        "group",
        vec![arr(), callback()],
        vec![ValType::Any],
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            use indexmap::IndexMap;
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let dispatch = classify_callback(&callback);
            let mut groups: IndexMap<String, Vec<Value>> = IndexMap::new();
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                groups.reserve(length);
                let receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver];
                for i in 0..length {
                    let Some(elem) = array_present_value_at(&arr, i) else {
                        continue;
                    };
                    invoke_args[0] = elem.clone();
                    invoke_args[1] = Value::I32(i as i32);
                    let key_value =
                        invoke_classified_callback(ctx, &callback, &dispatch, &invoke_args);
                    let key = crate::keys::with_property_key(&key_value, |key| key.to_owned());
                    groups
                        .entry(key)
                        .or_insert_with(|| Vec::with_capacity(4))
                        .push(elem);
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

    array_fn(
        vm,
        "groupToMap",
        vec![arr(), callback()],
        vec![ValType::Any],
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            use indexmap::IndexMap;
            let null = Value::Null;
            let callback = args.get(1).unwrap_or(&null);
            let dispatch = classify_callback(&callback);
            let mut groups: IndexMap<Value, Vec<Value>> = IndexMap::new();
            if let Some(arr) = array_of(args, 0) {
                let length = array_length(&arr);
                groups.reserve(length);
                let receiver = Value::Object(arr.clone());
                let mut invoke_args = [Value::Undefined, Value::I32(0), receiver];
                for i in 0..length {
                    let Some(elem) = array_present_value_at(&arr, i) else {
                        continue;
                    };
                    invoke_args[0] = elem.clone();
                    invoke_args[1] = Value::I32(i as i32);
                    let key = invoke_classified_callback(ctx, &callback, &dispatch, &invoke_args);
                    groups
                        .entry(key)
                        .or_insert_with(|| Vec::with_capacity(4))
                        .push(elem);
                }
            }
            // Build a JS Map with one entry per group.
            let mut map_im: IndexMap<Value, Value> = IndexMap::with_capacity(groups.len());
            for (k, v) in groups {
                map_im.insert(k, make_array(v));
            }
            let mut obj = Object::new();
            obj.kind = ObjectKind::Map(map_im);
            obj.properties
                .insert("size".into(), Value::I32(obj_map_len(&obj) as i32));
            // §24.1.3.10: `size` is a non-enumerable accessor — this cache of it
            // must not turn up in `for...in`.
            crate::object::track_nonenum_in(&mut obj, "size");
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
            let items = args.get(3..).unwrap_or(&[]);
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
                out.extend_from_slice(items);
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
        | Value::WeakRef(_) => true,
    }
}

fn obj_map_len(obj: &Object) -> usize {
    if let ObjectKind::Map(ref m) = obj.kind {
        m.len()
    } else {
        0
    }
}
