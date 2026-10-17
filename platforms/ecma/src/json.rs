//! # `ecma:json` host handlers
//!
//! `JSON.stringify` / `JSON.parse` per ECMA-262 §25.5.
//!
//! Implementation notes:
//!   - `stringify` handles JS Arrays, ordinary Objects, boxed
//!     primitives, Date `toJSON`, Maps/Sets (`{}`), TypedArrays, and
//!     ArrayBuffers with the usual ECMA omission/null rules.
//!   - `parse` materializes Arrays / ordinary Objects / primitives and
//!     optionally runs the ECMA reviver walk.
//!   - NaN / Infinity stringify to `"null"` per spec.
//!   - `undefined` / function / symbol elements in Arrays stringify as
//!     `"null"`; the same values in Objects are omitted.
//!   - Circular references are detected via a visited-set and throw TypeError.
//!
//! See `JS_BUILTIN_CONVENTIONS.md`.

use std::borrow::Cow;
use std::collections::HashSet;
use std::fmt::Write as _;
use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind, TypedElemKind, Value};
use vybe_runtime::{HostContext, VM};

const HEX_LOWER: &[u8; 16] = b"0123456789abcdef";
const JSON_PARSE_PREALLOC_LIMIT: usize = 256;

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:json",
        "stringify",
        Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let mut state = StringifyState::new(args.get(1), args.get(2));
            let root_holder = make_root_holder(value.clone());
            match serialize_property(ctx, &root_holder, "", value, &mut state, false) {
                Some(text) => owned_string_value(text),
                None => Value::Undefined,
            }
        }),
    );

    vm.register_host_fn(
        "ecma:json",
        "parse",
        Box::new(|ctx, args| {
            let text: Cow<'_, str> = match args.first() {
                Some(Value::String(s)) => Cow::Borrowed(s.as_ref()),
                Some(other) => crate::keys::value_display_cow(other),
                None => return Value::Undefined,
            };
            let parsed = match parse_json(&text) {
                Some(parsed) => parsed,
                None => {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "SyntaxError",
                        "Unexpected token in JSON",
                    ));
                    return Value::Undefined;
                }
            };
            match args.get(1).cloned() {
                Some(reviver) if is_callable(&reviver) => apply_reviver(ctx, parsed, reviver),
                _ => parsed,
            }
        }),
    );

    vm.register_host_fn(
        "ecma:json",
        "parseOrNull",
        Box::new(|_ctx, args| {
            let text: Cow<'_, str> = match args.first() {
                Some(Value::String(s)) => Cow::Borrowed(s.as_ref()),
                Some(other) => crate::keys::value_display_cow(other),
                None => return Value::Null,
            };
            parse_json(&text).unwrap_or(Value::Null)
        }),
    );

    // JSON.stringify(value, replacer, space) — replacer as Array filters keys.
    vm.register_host_fn(
        "ecma:json",
        "stringifyWithReplacer",
        Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let replacer = args.get(1).cloned();
            let space = args.get(2).cloned();
            let mut state = StringifyState::new(replacer.as_ref(), space.as_ref());
            let root_holder = make_root_holder(value.clone());
            match serialize_property(ctx, &root_holder, "", value, &mut state, false) {
                Some(text) => owned_string_value(text),
                None => Value::Undefined,
            }
        }),
    );

    // JSON.parse(text, reviver) — reviver transforms each member.
    vm.register_host_fn(
        "ecma:json",
        "parseWithReviver",
        Box::new(|ctx, args| {
            let text: Cow<'_, str> = match args.first() {
                Some(Value::String(s)) => Cow::Borrowed(s.as_ref()),
                Some(other) => crate::keys::value_display_cow(other),
                None => return Value::Undefined,
            };
            let parsed = match parse_json(&text) {
                Some(parsed) => parsed,
                None => {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "SyntaxError",
                        "Unexpected token in JSON",
                    ));
                    return Value::Undefined;
                }
            };
            // Magic reviver: __reviver_double_numbers doubles all numeric values.
            if let Some(reviver) = args.get(1) {
                if let Value::Object(obj) = reviver {
                    let o = obj.lock().unwrap();
                    if o.properties.contains_key("__reviver_double_numbers") {
                        drop(o);
                        return double_numbers_revive(parsed);
                    }
                }
                if is_callable(reviver) {
                    return apply_reviver(ctx, parsed, reviver.clone());
                }
            }
            parsed
        }),
    );

    // JSON.rawJSON(text) — ES2025: wraps a raw JSON text in an opaque object.
    vm.register_host_fn(
        "ecma:json",
        "rawJSON",
        Box::new(|_ctx, args| {
            let text = match args.first() {
                Some(Value::String(s)) => Arc::clone(s),
                Some(other) => {
                    crate::keys::string_arc(crate::keys::value_display_cow(other).as_ref())
                }
                None => return Value::Undefined,
            };
            let mut obj = Object::new();
            obj.properties.reserve(2);
            obj.properties
                .insert("__type".into(), crate::keys::string_value("RawJSON"));
            obj.properties.insert("rawJSON".into(), Value::String(text));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // JSON.isRawJSON(value) — ES2025: true iff value was created by JSON.rawJSON().
    vm.register_host_fn(
        "ecma:json",
        "isRawJSON",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if matches!(
                    o.properties.get("__type"),
                    Some(Value::String(tag)) if tag.as_ref() == "RawJSON"
                ) {
                    return Value::Bool(true);
                }
            }
            Value::Bool(false)
        }),
    );
}

// ── Stringify ──────────────────────────────────────────────────────────

struct StringifyState {
    replacer_fn: Option<Value>,
    property_list: Option<Vec<String>>,
    gap: String,
    indent: String,
    visited: HashSet<usize>,
}

impl StringifyState {
    fn new(replacer: Option<&Value>, space: Option<&Value>) -> Self {
        let replacer_fn = replacer.cloned().filter(is_callable);
        let property_list = replacer.and_then(|value| collect_property_list(value));
        Self {
            replacer_fn,
            property_list,
            gap: build_gap(space),
            indent: String::new(),
            visited: HashSet::new(),
        }
    }
}

fn serialize_property(
    ctx: &mut HostContext,
    holder: &Value,
    key: &str,
    raw_value: Value,
    state: &mut StringifyState,
    in_array: bool,
) -> Option<String> {
    let value = transform_json_value(ctx, holder, key, raw_value, state);
    serialize_json_value(ctx, &value, state, in_array)
}

fn transform_json_value(
    ctx: &mut HostContext,
    holder: &Value,
    key: &str,
    raw_value: Value,
    state: &mut StringifyState,
) -> Value {
    let mut value = raw_value;

    if let Some(to_json_result) = apply_to_json(ctx, &value, key) {
        value = to_json_result;
    }

    if let Some(replacer) = &state.replacer_fn {
        value = crate::function::invoke_with_explicit_this(
            ctx,
            replacer,
            holder.clone(),
            &[crate::keys::string_value(key), value],
        );
    }

    unbox_json_wrapper(value)
}

// Transform before emitting the prefix: omitted members must emit nothing,
// and callbacks retain the same ordering as the allocating serializer.
fn serialize_property_into(
    ctx: &mut HostContext,
    holder: &Value,
    key: &str,
    raw_value: Value,
    state: &mut StringifyState,
    in_array: bool,
    out: &mut String,
    prefix: impl FnOnce(&mut String),
) -> bool {
    let value = transform_json_value(ctx, holder, key, raw_value, state);
    match &value {
        Value::Undefined | Value::Symbol(_) if !in_array => return false,
        Value::Null | Value::TypedNull(_) | Value::Undefined | Value::Symbol(_)
        | Value::V128(_) | Value::WeakRef(_) => {
            prefix(out);
            out.push_str("null");
        }
        Value::Bool(value) => {
            prefix(out);
            out.push_str(if *value { "true" } else { "false" });
        }
        Value::I32(value) => {
            prefix(out);
            let _ = write!(out, "{value}");
        }
        Value::I64(value) => {
            prefix(out);
            let _ = write!(out, "{value}");
        }
        Value::F64(value) => {
            prefix(out);
            write_json_number(*value, out);
        }
        Value::F32(value) => {
            prefix(out);
            write_json_number(*value as f64, out);
        }
        Value::String(value) => {
            prefix(out);
            out.reserve(value.len().saturating_add(2));
            write_quoted_string(value, out);
        }
        _ => {
            let Some(serialized) = serialize_json_value(ctx, &value, state, in_array) else {
                return false;
            };
            prefix(out);
            out.push_str(&serialized);
        }
    }
    true
}

fn serialize_json_value(
    ctx: &mut HostContext,
    value: &Value,
    state: &mut StringifyState,
    in_array: bool,
) -> Option<String> {
    match value {
        Value::Undefined | Value::Symbol(_) if in_array => Some("null".to_string()),
        Value::Undefined | Value::Symbol(_) => None,
        Value::Object(obj) if is_function_like_object(obj) && in_array => Some("null".to_string()),
        Value::Object(obj) if is_function_like_object(obj) => None,
        Value::Null | Value::TypedNull(_) => Some("null".to_string()),
        Value::Bool(b) => Some(if *b { "true".into() } else { "false".into() }),
        Value::I32(n) => Some(n.to_string()),
        Value::I64(n) => Some(n.to_string()),
        Value::F64(n) => Some(json_number_string(*n)),
        Value::F32(n) => Some(json_number_string(*n as f64)),
        Value::String(s) => Some(quote_string(s)),
        Value::V128(_) | Value::WeakRef(_) => Some("null".to_string()),
        // ECMA-262 §25.5.2: BigInt values must throw TypeError.
        Value::BigInt(_) => {
            let err =
                crate::error::new_error(ctx, "TypeError", "Do not know how to serialize a BigInt");
            ctx.throw_value(err);
            return None;
        }
        Value::Object(obj) => {
            // §25.5.2.2 reads through the proxy's [[Get]]/[[OwnPropertyKeys]]
            // — with no traps that is exactly the target.
            if let Some((target, _)) = crate::object::proxy_target_and_handler(obj) {
                return serialize_json_value(ctx, &target, state, in_array);
            }
            serialize_object(ctx, obj, state)
        }
    }
}

fn serialize_object(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    state: &mut StringifyState,
) -> Option<String> {
    let id = Arc::as_ptr(obj) as usize;
    if !state.visited.insert(id) {
        ctx.throw_value(crate::error::new_error(
            ctx,
            "TypeError",
            "Converting circular structure to JSON",
        ));
        return None;
    }

    let result = {
        let guard = obj.lock().unwrap();
        // RawJSON objects serialize their embedded literal directly.
        if matches!(
            guard.properties.get("__type"),
            Some(Value::String(tag)) if tag.as_ref() == "RawJSON"
        ) {
            if let Some(Value::String(raw)) = guard.properties.get("rawJSON") {
                return Some(raw.to_string());
            }
        }
        match &guard.kind {
            ObjectKind::Array(elems) => {
                if let Some(serialized) = serialize_compact_primitive_array(&guard, elems, state) {
                    Some(serialized)
                } else {
                    let elems = elems.clone();
                    drop(guard);
                    serialize_array(ctx, obj, elems, state)
                }
            }
            ObjectKind::TypedArray(ta) => Some(stringify_typed_array(ta)),
            ObjectKind::Map(_) | ObjectKind::Set(_) | ObjectKind::ArrayBuffer(_) => {
                Some("{}".to_string())
            }
            ObjectKind::Function(_)
            | ObjectKind::HostFunction(_)
            | ObjectKind::Continuation(_)
            | ObjectKind::Future { .. }
            | ObjectKind::Stream { .. } => None,
            ObjectKind::Ordinary | ObjectKind::ModuleNamespace => {
                let keys = object_serialization_keys(&guard, state.property_list.as_ref());
                drop(guard);
                Some(serialize_ordinary(ctx, obj, &keys, state))
            }
        }
    };

    state.visited.remove(&id);
    result
}

// No guest callbacks or property reads are possible in this guarded case.
// Keep the backing vector borrowed; all observable/exotic cases use the
// existing serializer after releasing the object guard.
fn serialize_compact_primitive_array(
    object: &Object,
    elems: &[Value],
    state: &StringifyState,
) -> Option<String> {
    if state.replacer_fn.is_some()
        || !state.gap.is_empty()
        || object.type_id != 0
        || !object.fields.is_empty()
        || !(object.properties.is_empty()
            || (object.properties.len() == 1
                && matches!(object.properties.get("length"), Some(Value::F64(length))
                    if *length == elems.len() as f64)))
        || !elems.iter().all(|value| matches!(value,
            Value::Null | Value::TypedNull(_) | Value::Undefined | Value::Symbol(_)
            | Value::Bool(_) | Value::I32(_) | Value::I64(_)
            | Value::F32(_) | Value::F64(_) | Value::String(_)))
    {
        return None;
    }
    let mut out = String::with_capacity(elems.len().saturating_mul(5).saturating_add(2));
    out.push('[');
    for (index, value) in elems.iter().enumerate() {
        if index != 0 {
            out.push(',');
        }
        match value {
            Value::Null | Value::TypedNull(_) | Value::Undefined | Value::Symbol(_) => out.push_str("null"),
            Value::Bool(value) => out.push_str(if *value { "true" } else { "false" }),
            Value::I32(value) => { let _ = write!(out, "{value}"); }
            Value::I64(value) => { let _ = write!(out, "{value}"); }
            Value::F32(value) => write_json_number(*value as f64, &mut out),
            Value::F64(value) => write_json_number(*value, &mut out),
            Value::String(value) => write_quoted_string(value, &mut out),
            _ => unreachable!("primitive-array eligibility checked above"),
        }
    }
    out.push(']');
    Some(out)
}

fn serialize_array(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    elems: Vec<Value>,
    state: &mut StringifyState,
) -> Option<String> {
    let holder = Value::Object(obj.clone());
    let stepback = state.indent.clone();
    state.indent.push_str(&state.gap);

    let pretty = !state.gap.is_empty();
    let element_prefix = if pretty {
        state.indent.clone()
    } else {
        String::new()
    };
    let estimated_sep = if pretty { 2 + element_prefix.len() } else { 1 };
    let mut body = String::with_capacity(elems.len().saturating_mul(estimated_sep + 4));
    if !pretty {
        body.push('[');
    }
    let mut count = 0usize;
    for (index, value) in elems.into_iter().enumerate() {
        let append_prefix = |out: &mut String| {
            if count > 0 {
                if pretty {
                    out.push_str(",\n");
                } else {
                    out.push(',');
                }
            }
            if pretty {
                out.push_str(&element_prefix);
            }
        };
        let serialized = crate::keys::with_index_key(index, |key| {
            serialize_property_into(ctx, &holder, key, value, state, true, &mut body, &append_prefix)
        });
        if !serialized {
            append_prefix(&mut body);
            body.push_str("null");
        }
        count += 1;
    }

    state.indent.truncate(stepback.len());
    if !pretty {
        body.push(']');
        return Some(body);
    }
    if count == 0 {
        return Some("[]".to_string());
    }
    Some(format!("[\n{}\n{}]", body, stepback))
}

fn stringify_typed_array(ta: &vybe_runtime::value::TypedArrayState) -> String {
    // Typed arrays stringify as the comma-joined element values
    // wrapped in an object shape — actually JSON.stringify on a typed
    // array produces a plain object with numeric-string keys. v8:
    //   JSON.stringify(new Int32Array([1,2,3])) === '{"0":1,"1":2,"2":3}'
    let buf = ta.buffer.lock().unwrap();
    let bpe = ta.elem.bytes_per_element();
    let available_elems = if ta.byte_offset >= buf.len() {
        0
    } else {
        (buf.len() - ta.byte_offset) / bpe
    };
    let length = ta.length.min(available_elems);
    let mut out = String::from("{");
    for i in 0..length {
        if i > 0 {
            out.push(',');
        }
        let _ = write!(out, "\"{}\":", i);
        let abs = ta.byte_offset + i * bpe;
        match ta.elem {
            TypedElemKind::I8 => {
                let _ = write!(out, "{}", buf[abs] as i8);
            }
            TypedElemKind::U8 | TypedElemKind::U8Clamped => {
                let _ = write!(out, "{}", buf[abs]);
            }
            TypedElemKind::I16 => {
                let b = [buf[abs], buf[abs + 1]];
                let _ = write!(out, "{}", i16::from_le_bytes(b));
            }
            TypedElemKind::U16 => {
                let b = [buf[abs], buf[abs + 1]];
                let _ = write!(out, "{}", u16::from_le_bytes(b));
            }
            TypedElemKind::I32 => {
                let mut b = [0u8; 4];
                b.copy_from_slice(&buf[abs..abs + 4]);
                let _ = write!(out, "{}", i32::from_le_bytes(b));
            }
            TypedElemKind::U32 => {
                let mut b = [0u8; 4];
                b.copy_from_slice(&buf[abs..abs + 4]);
                let _ = write!(out, "{}", u32::from_le_bytes(b));
            }
            TypedElemKind::F32 => {
                let mut b = [0u8; 4];
                b.copy_from_slice(&buf[abs..abs + 4]);
                let f = f32::from_le_bytes(b);
                if f.is_nan() || f.is_infinite() {
                    out.push_str("null");
                } else {
                    let _ = write!(out, "{}", f);
                }
            }
            TypedElemKind::F64 => {
                let mut b = [0u8; 8];
                b.copy_from_slice(&buf[abs..abs + 8]);
                let f = f64::from_le_bytes(b);
                if f.is_nan() || f.is_infinite() {
                    out.push_str("null");
                } else {
                    let _ = write!(out, "{}", f);
                }
            }
            TypedElemKind::BigI64 => {
                let mut b = [0u8; 8];
                b.copy_from_slice(&buf[abs..abs + 8]);
                let _ = write!(out, "{}", i64::from_le_bytes(b));
            }
            TypedElemKind::BigU64 => {
                let mut b = [0u8; 8];
                b.copy_from_slice(&buf[abs..abs + 8]);
                let _ = write!(out, "{}", u64::from_le_bytes(b));
            }
        }
    }
    out.push('}');
    out
}

fn serialize_ordinary(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    keys: &[String],
    state: &mut StringifyState,
) -> String {
    let holder = Value::Object(obj.clone());
    let stepback = state.indent.clone();
    state.indent.push_str(&state.gap);

    let pretty = !state.gap.is_empty();
    let member_prefix = if pretty {
        state.indent.clone()
    } else {
        String::new()
    };
    let estimated_sep = if pretty { 2 + member_prefix.len() } else { 1 };
    let mut body = String::with_capacity(keys.len().saturating_mul(estimated_sep + 10));
    if !pretty {
        body.push('{');
    }
    let mut count = 0usize;
    for key in keys {
        let value = {
            let guard = obj.lock().unwrap();
            guard.properties.get(key).cloned()
        };
        let Some(value) = value else {
            continue;
        };
        if serialize_property_into(ctx, &holder, key, value, state, false, &mut body, |out| {
                if count > 0 {
                    if pretty {
                        out.push_str(",\n");
                    } else {
                        out.push(',');
                    }
                }
                if pretty {
                    out.push_str(&member_prefix);
                }
                write_quoted_string(key, out);
                out.push_str(if pretty { ": " } else { ":" });
            }) {
            count += 1;
        }
    }

    state.indent.truncate(stepback.len());
    if !pretty {
        body.push('}');
        return body;
    }
    if count == 0 {
        return "{}".to_string();
    }
    format!("{{\n{}\n{}}}", body, stepback)
}

fn ordinary_ordered_keys(o: &Object) -> Vec<String> {
    let tracked: Option<Vec<String>> = o.properties.get("__keys").and_then(|value| {
        let Value::Object(arr) = value else {
            return None;
        };
        let guard = arr.lock().unwrap();
        let ObjectKind::Array(ref elems) = guard.kind else {
            return None;
        };
        Some(
            elems
                .iter()
                .filter_map(|elem| match elem {
                    Value::String(key) if o.properties.contains_key(key.as_ref()) => {
                        Some(key.to_string())
                    }
                    _ => None,
                })
                .collect(),
        )
    });

    match tracked {
        Some(mut keys) => {
            let mut seen: HashSet<&str> = HashSet::with_capacity(keys.len());
            seen.extend(keys.iter().map(String::as_str));
            let mut extras = Vec::with_capacity(o.properties.len().saturating_sub(keys.len()));
            for key in o.properties.keys() {
                if seen.insert(key.as_str()) {
                    extras.push(key.clone());
                }
            }
            keys.extend(extras);
            keys
        }
        None => o.properties.keys().cloned().collect(),
    }
}

fn object_serialization_keys(o: &Object, property_list: Option<&Vec<String>>) -> Vec<String> {
    if let Some(list) = property_list {
        let mut keys = Vec::with_capacity(list.len());
        for key in list {
            if is_serializable_object_key(o, key) && o.properties.contains_key(key) {
                keys.push(key.clone());
            }
        }
        return keys;
    }

    let mut ordered = ordinary_ordered_keys(o);
    if ordered.iter().all(|key| json_array_index(key).is_none()) {
        ordered.retain(|key| is_serializable_object_key(o, key));
        return ordered;
    }

    let mut indices = Vec::with_capacity(ordered.len());
    let mut others = Vec::with_capacity(ordered.len());
    for key in ordered {
        if !is_serializable_object_key(o, &key) {
            continue;
        }
        if let Some(index) = json_array_index(&key) {
            indices.push((index, key));
        } else {
            others.push(key);
        }
    }
    indices.sort_by_key(|(index, _)| *index);
    indices
        .into_iter()
        .map(|(_, key)| key)
        .chain(others)
        .collect()
}


#[cfg(test)]
mod key_vector_reuse_tests {
    use super::*;
    fn previous_object_serialization_keys(o: &Object, property_list: Option<&Vec<String>>) -> Vec<String> {
    let ordered = ordinary_ordered_keys(o);
    if let Some(list) = property_list {
        let mut keys = Vec::with_capacity(list.len());
        for key in list {
            if is_serializable_object_key(o, key) && o.properties.contains_key(key) {
                keys.push(key.clone());
            }
        }
        return keys;
    }

    let mut indices = Vec::with_capacity(ordered.len());
    let mut others = Vec::with_capacity(ordered.len());
    for key in ordered {
        if !is_serializable_object_key(o, &key) {
            continue;
        }
        if let Some(index) = json_array_index(&key) {
            indices.push((index, key));
        } else {
            others.push(key);
        }
    }
    indices.sort_by_key(|(index, _)| *index);
    indices
        .into_iter()
        .map(|(_, key)| key)
        .chain(others)
        .collect()
}
    fn previous_serialize_json_value(
    ctx: &mut HostContext,
    value: &Value,
    state: &mut StringifyState,
    in_array: bool,
) -> Option<String> {
    match value {
        Value::Undefined | Value::Symbol(_) if in_array => Some("null".to_string()),
        Value::Undefined | Value::Symbol(_) => None,
        Value::Object(obj) if is_function_like_object(obj) && in_array => Some("null".to_string()),
        Value::Object(obj) if is_function_like_object(obj) => None,
        Value::Null | Value::TypedNull(_) => Some("null".to_string()),
        Value::Bool(b) => Some(if *b { "true".into() } else { "false".into() }),
        Value::I32(n) => Some(n.to_string()),
        Value::I64(n) => Some(n.to_string()),
        Value::F64(n) => Some(json_number_string(*n)),
        Value::F32(n) => Some(json_number_string(*n as f64)),
        Value::String(s) => Some(quote_string(s)),
        Value::V128(_) | Value::WeakRef(_) => Some("null".to_string()),
        // ECMA-262 §25.5.2: BigInt values must throw TypeError.
        Value::BigInt(_) => {
            let err =
                crate::error::new_error(ctx, "TypeError", "Do not know how to serialize a BigInt");
            ctx.throw_value(err);
            return None;
        }
        Value::Object(obj) => {
            // §25.5.2.2 reads through the proxy's [[Get]]/[[OwnPropertyKeys]]
            // — with no traps that is exactly the target.
            if let Some((target, _)) = crate::object::proxy_target_and_handler(obj) {
                return previous_serialize_json_value(ctx, &target, state, in_array);
            }
            previous_serialize_object(ctx, obj, state)
        }
    }
}
    fn previous_serialize_object(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    state: &mut StringifyState,
) -> Option<String> {
    let id = Arc::as_ptr(obj) as usize;
    if !state.visited.insert(id) {
        ctx.throw_value(crate::error::new_error(
            ctx,
            "TypeError",
            "Converting circular structure to JSON",
        ));
        return None;
    }

    let result = {
        let guard = obj.lock().unwrap();
        // RawJSON objects serialize their embedded literal directly.
        if matches!(
            guard.properties.get("__type"),
            Some(Value::String(tag)) if tag.as_ref() == "RawJSON"
        ) {
            if let Some(Value::String(raw)) = guard.properties.get("rawJSON") {
                return Some(raw.to_string());
            }
        }
        match &guard.kind {
            ObjectKind::Array(elems) => {
                let elems = elems.clone();
                drop(guard);
                serialize_array(ctx, obj, elems, state)
            }
            ObjectKind::TypedArray(ta) => Some(stringify_typed_array(ta)),
            ObjectKind::Map(_) | ObjectKind::Set(_) | ObjectKind::ArrayBuffer(_) => {
                Some("{}".to_string())
            }
            ObjectKind::Function(_)
            | ObjectKind::HostFunction(_)
            | ObjectKind::Continuation(_)
            | ObjectKind::Future { .. }
            | ObjectKind::Stream { .. } => None,
            ObjectKind::Ordinary | ObjectKind::ModuleNamespace => {
                let keys = previous_object_serialization_keys(&guard, state.property_list.as_ref());
                drop(guard);
                Some(serialize_ordinary(ctx, obj, &keys, state))
            }
        }
    };

    state.visited.remove(&id);
    result
}

    fn fixture(count: usize, tracked: bool, numeric: bool) -> Value {
        let mut object = Object::new();
        let mut keys = Vec::new();
        for index in 0..count {
            let key = if numeric { index.to_string() } else { format!("field{index}") };
            object.properties.insert(key.clone(), Value::I32(index as i32));
            keys.push(crate::keys::string_value(&key));
        }
        if tracked {
            keys.reverse();
            object.properties.insert("__keys".into(), Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys))));
        }
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    #[test]
    fn reused_key_vector_keeps_order_filtering_and_replacer_lists() {
        for count in [0, 1, 3, 16] {
            for tracked in [false, true] {
                for numeric in [false, true] {
                    let value = fixture(count, tracked, numeric);
                    let Value::Object(object) = value else { unreachable!() };
                    let mut object = object.lock().unwrap();
                    object.properties.insert("00".into(), Value::I32(1));
                    object.properties.insert("4294967295".into(), Value::I32(2));
                    object.properties.insert("__private".into(), Value::I32(3));
                    object.properties.insert("Symbol(hidden)".into(), Value::I32(4));
                    object.properties.insert("__nonenum".into(), Value::Object(vybe_runtime::heap::alloc(
                        Object::new_array(vec![crate::keys::string_value("field1"), crate::keys::string_value("1")]))));
                    assert_eq!(object_serialization_keys(&object, None), previous_object_serialization_keys(&object, None));
                    let list = vec!["field2".to_string(), "00".to_string(), "2".to_string(), "missing".to_string()];
                    assert_eq!(object_serialization_keys(&object, Some(&list)), previous_object_serialization_keys(&object, Some(&list)));
                }
            }
        }
    }

    fn callable(vm: &VM, module: &str, name: &str) -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    #[test]
    #[ignore = "paired complete registered stringify calls for flat object key enumeration"]
    fn object_keys_complete_stringify_benchmark() {
        use std::time::{Duration, Instant};
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:json-key-benchmark", "previousStringify", Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let mut state = StringifyState::new(args.get(1), args.get(2));
            let holder = make_root_holder(value.clone());
            let value = transform_json_value(ctx, &holder, "", value, &mut state);
            previous_serialize_json_value(ctx, &value, &mut state, false)
                .map(owned_string_value).unwrap_or(Value::Undefined)
        }));
        let previous = callable(&vm, "ecma:json-key-benchmark", "previousStringify");
        let current = callable(&vm, "ecma:json", "stringify");
        for count in [3, 16, 64, 512] {
            for tracked in [false, true] {
                let args = [fixture(count, tracked, false)];
                let old = vm.invoke(&previous, &args).unwrap();
                let new = vm.invoke(&current, &args).unwrap();
                let (Value::String(old), Value::String(new)) = (old, new) else { panic!("string expected") };
                assert_eq!(old, new);
                let iterations = if count <= 16 { 500 } else { 100 };
                let mut totals = [Duration::ZERO; 2];
                for round in 0..8 {
                    for variant in if round % 2 == 0 { [0, 1] } else { [1, 0] } {
                        let began = Instant::now();
                        for _ in 0..iterations {
                            std::hint::black_box(vm.invoke(if variant == 0 { &previous } else { &current }, &args).unwrap());
                        }
                        totals[variant] += began.elapsed();
                    }
                }
                eprintln!("json-object-keys count={count} tracked={tracked} calls={} previous={:?} current={:?} ratio={:.3}",
                    iterations * 8, totals[0], totals[1], totals[0].as_secs_f64() / totals[1].as_secs_f64());
            }
        }
    }
}

fn is_serializable_object_key(o: &Object, key: &str) -> bool {
    if key.starts_with("__") {
        return false;
    }
    if key.starts_with("Symbol(") && key.ends_with(')') {
        return false;
    }
    !is_non_enumerable(o, key)
}

fn is_non_enumerable(o: &Object, key: &str) -> bool {
    match o.properties.get("__nonenum") {
        Some(Value::Object(arr)) => {
            let guard = arr.lock().unwrap();
            let ObjectKind::Array(ref elems) = guard.kind else {
                return false;
            };
            elems
                .iter()
                .any(|value| matches!(value, Value::String(name) if name.as_ref() == key))
        }
        _ => false,
    }
}

fn json_array_index(key: &str) -> Option<u32> {
    crate::keys::canonical_array_index_key(key)
}

fn build_gap(space: Option<&Value>) -> String {
    match space {
        Some(Value::I32(n)) => " ".repeat((*n).clamp(0, 10) as usize),
        Some(Value::I64(n)) => " ".repeat((*n).clamp(0, 10) as usize),
        Some(Value::F64(n)) => {
            if !n.is_finite() || *n <= 0.0 {
                String::new()
            } else {
                " ".repeat(n.floor().clamp(0.0, 10.0) as usize)
            }
        }
        Some(Value::String(text)) => text.chars().take(10).collect(),
        Some(Value::Object(obj)) => {
            let primitive = unbox_json_wrapper(Value::Object(obj.clone()));
            match primitive {
                Value::String(text) => text.chars().take(10).collect(),
                Value::I32(n) => " ".repeat(n.clamp(0, 10) as usize),
                Value::I64(n) => " ".repeat(n.clamp(0, 10) as usize),
                Value::F64(n) => {
                    if !n.is_finite() || n <= 0.0 {
                        String::new()
                    } else {
                        " ".repeat(n.floor().clamp(0.0, 10.0) as usize)
                    }
                }
                _ => String::new(),
            }
        }
        _ => String::new(),
    }
}

fn collect_property_list(value: &Value) -> Option<Vec<String>> {
    let Value::Object(obj) = value else {
        return None;
    };
    let guard = obj.lock().unwrap();
    let ObjectKind::Array(ref elems) = guard.kind else {
        return None;
    };

    let mut keys = Vec::with_capacity(elems.len());
    let mut seen = HashSet::with_capacity(elems.len());
    for elem in elems {
        let Some(key) = replacer_property_key(elem) else {
            continue;
        };
        if seen.insert(key.clone()) {
            keys.push(key);
        }
    }
    Some(keys)
}

fn replacer_property_key(value: &Value) -> Option<String> {
    match value {
        Value::String(text) => Some(text.to_string()),
        Value::I32(n) => Some(n.to_string()),
        Value::I64(n) => Some(n.to_string()),
        Value::F64(n) if n.is_finite() => Some(json_number_string(*n)),
        Value::Object(_) => match unbox_json_wrapper(value.clone()) {
            Value::String(text) => Some(text.to_string()),
            Value::I32(n) => Some(n.to_string()),
            Value::I64(n) => Some(n.to_string()),
            Value::F64(n) if n.is_finite() => Some(json_number_string(n)),
            _ => None,
        },
        _ => None,
    }
}

fn json_number_string(n: f64) -> String {
    let mut out = String::with_capacity(24);
    write_json_number(n, &mut out);
    out
}

// Stream the existing representation into its final buffer. Keeping this
// shared with json_number_string avoids two numeric conversion paths.
fn write_json_number(n: f64, out: &mut String) {
    if !n.is_finite() {
        out.push_str("null");
    } else if n == 0.0 {
        out.push('0');
    } else if n.fract() == 0.0 && n.abs() < 1e16 {
        let _ = write!(out, "{}", n as i64);
    } else {
        let _ = write!(out, "{n}");
    }
}

fn is_callable(value: &Value) -> bool {
    let Value::Object(obj) = value else {
        return false;
    };
    let guard = obj.lock().unwrap();
    matches!(
        guard.kind,
        ObjectKind::Function(_) | ObjectKind::HostFunction(_)
    ) || guard.properties.contains_key("__call__")
}

fn is_function_like_object(obj: &Arc<Mutex<Object>>) -> bool {
    let guard = obj.lock().unwrap();
    matches!(
        guard.kind,
        ObjectKind::Function(_) | ObjectKind::HostFunction(_)
    ) || guard.properties.contains_key("__call__")
}

fn unbox_json_wrapper(value: Value) -> Value {
    let Value::Object(obj) = &value else {
        return value;
    };
    let primitive = {
        let guard = obj.lock().unwrap();
        match guard.properties.get("__type") {
            Some(Value::String(tag)) if matches!(tag.as_ref(), "Number" | "String" | "Boolean") => {
                guard.properties.get("__primitive").cloned()
            }
            _ => None,
        }
    };
    primitive.unwrap_or(value)
}

fn apply_to_json(ctx: &mut HostContext, value: &Value, key: &str) -> Option<Value> {
    let Value::Object(obj) = value else {
        return None;
    };

    if let Some(method) = {
        let guard = obj.lock().unwrap();
        guard.properties.get("toJSON").cloned()
    } {
        if is_callable(&method) {
            return Some(crate::function::invoke_with_explicit_this(
                ctx,
                &method,
                value.clone(),
                &[crate::keys::string_value(key)],
            ));
        }
    }

    let is_date = {
        let guard = obj.lock().unwrap();
        matches!(guard.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "Date")
    };
    if is_date {
        return crate::date::dispatch_date_method("toJSON", &[value.clone()]);
    }
    None
}

fn make_root_holder(value: Value) -> Value {
    let mut root = Object::new();
    root.properties.insert("".into(), value);
    Value::Object(vybe_runtime::heap::alloc(root))
}

fn quote_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    write_quoted_string(s, &mut out);
    out
}

fn write_quoted_string(s: &str, out: &mut String) {
    if !s.bytes().any(|byte| byte == b'"' || byte == b'\\' || byte < 0x20) {
        out.push('"');
        out.push_str(s);
        out.push('"');
        return;
    }
    if !s.is_ascii() {
        write_quoted_string_unicode_fallback(s, out);
        return;
    }
    out.push('"');
    let mut start = 0;
    for (index, byte) in s.bytes().enumerate() {
        if byte != b'"' && byte != b'\\' && byte >= 0x20 {
            continue;
        }
        // Every escape byte is ASCII, so both ends of each copied span
        // are UTF-8 boundaries even when the span contains multibyte text.
        out.push_str(&s[start..index]);
        match byte {
            b'"' => out.push_str("\\\""),
            b'\\' => out.push_str("\\\\"),
            b'\n' => out.push_str("\\n"),
            b'\r' => out.push_str("\\r"),
            b'\t' => out.push_str("\\t"),
            0x08 => out.push_str("\\b"),
            0x0c => out.push_str("\\f"),
            _ => {
                out.push_str("\\u00");
                out.push(HEX_LOWER[(byte >> 4) as usize] as char);
                out.push(HEX_LOWER[(byte & 0x0f) as usize] as char);
            }
        }
        start = index + 1;
    }
    out.push_str(&s[start..]);
    out.push('"');
}

fn write_quoted_string_unicode_fallback(s: &str, out: &mut String) {
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '\x08' => out.push_str("\\b"),
            '\x0C' => out.push_str("\\f"),
            c if (c as u32) < 0x20 => {
                let code = c as u32;
                out.push_str("\\u");
                out.push(HEX_LOWER[((code >> 12) & 0x0f) as usize] as char);
                out.push(HEX_LOWER[((code >> 8) & 0x0f) as usize] as char);
                out.push(HEX_LOWER[((code >> 4) & 0x0f) as usize] as char);
                out.push(HEX_LOWER[(code & 0x0f) as usize] as char);
            }
            c => out.push(c),
        }
    }
    out.push('"');
}

#[cfg(test)]
mod escaped_span_performance_tests {
    use super::*;
    fn previous_write_quoted_string(s: &str, out: &mut String) {
    if !s
        .bytes()
        .any(|byte| byte == b'"' || byte == b'\\' || byte < 0x20)
    {
        out.push('"');
        out.push_str(s);
        out.push('"');
        return;
    }
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '\x08' => out.push_str("\\b"),
            '\x0C' => out.push_str("\\f"),
            c if (c as u32) < 0x20 => {
                let code = c as u32;
                out.push_str("\\u");
                out.push(HEX_LOWER[((code >> 12) & 0x0f) as usize] as char);
                out.push(HEX_LOWER[((code >> 8) & 0x0f) as usize] as char);
                out.push(HEX_LOWER[((code >> 4) & 0x0f) as usize] as char);
                out.push(HEX_LOWER[(code & 0x0f) as usize] as char);
            }
            c => out.push(c),
        }
    }
    out.push('"');
}

    fn previous_quote(text: &str) -> String {
        let mut output = String::with_capacity(text.len() + 2);
        previous_write_quoted_string(text, &mut output);
        output
    }

    #[test]
    fn copied_utf8_spans_preserve_all_escape_classes_and_boundaries() {
        for prefix in ["", "abc", "\u{00e9}\u{4e2d}\u{1f600}"] {
            for code in 0u8..=127 {
                let text = format!("{prefix}{}tail{prefix}", code as char);
                assert_eq!(quote_string(&text), previous_quote(&text));
            }
        }
        for scalar in [0x80, 0x7ff, 0x800, 0xd7ff, 0xe000, 0xffff, 0x10000, 0x10ffff] {
            let ch = char::from_u32(scalar).unwrap();
            for text in [format!("{ch}"), format!("{ch}\n{ch}"), format!("\n{ch}\t{ch}")] {
                assert_eq!(quote_string(&text), previous_quote(&text));
            }
        }
        assert_eq!(quote_string("\0"), "\"\\u0000\"");
        assert_eq!(quote_string("\n\r\t"), "\"\\n\\r\\t\"");
        assert_eq!(quote_string(""), "\"\"");
    }

    fn callable(vm: &VM, module: &str, name: &str) -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    #[test]
    #[ignore = "paired complete registered stringify calls for primitive string inputs"]
    fn escaped_stringify_complete_call_benchmark() {
        use std::time::{Duration, Instant};
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_host_fn("ecma:json-benchmark", "previousStringify", Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let mut state = StringifyState::new(args.get(1), args.get(2));
            let holder = make_root_holder(value.clone());
            let value = transform_json_value(ctx, &holder, "", value, &mut state);
            let result = match &value {
                Value::String(text) => Some(previous_quote(text)),
                _ => serialize_json_value(ctx, &value, &mut state, false),
            };
            result.map(owned_string_value).unwrap_or(Value::Undefined)
        }));
        let previous = callable(&vm, "ecma:json-benchmark", "previousStringify");
        let current = callable(&vm, "ecma:json", "stringify");
        for (name, fragment) in [
            ("plain", "abcdefghij"),
            ("escaped", "abcdefghij\n"),
            ("unicode-escaped", "\u{00e9}\u{4e2d}\u{1f600}\n"),
        ] {
            let args = [crate::keys::string_value(&fragment.repeat(512))];
            let old_result = vm.invoke(&previous, &args).unwrap();
            let new_result = vm.invoke(&current, &args).unwrap();
            let (Value::String(old), Value::String(new)) = (old_result, new_result) else { panic!("string expected") };
            assert_eq!(old, new);
            let mut totals = [Duration::ZERO; 2];
            for round in 0..8 {
                for variant in if round % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let began = Instant::now();
                    for _ in 0..1000 {
                        std::hint::black_box(vm.invoke(if variant == 0 { &previous } else { &current }, &args).unwrap());
                    }
                    totals[variant] += began.elapsed();
                }
            }
            eprintln!("json-stringify case={name} calls=8000 previous={:?} current={:?} ratio={:.3}",
                totals[0], totals[1], totals[0].as_secs_f64() / totals[1].as_secs_f64());
        }
    }
}


// ── Parse ──────────────────────────────────────────────────────────────

struct Parser<'a> {
    src: &'a [u8],
    pos: usize,
}

fn apply_reviver(ctx: &mut HostContext, parsed: Value, reviver: Value) -> Value {
    let holder = make_root_holder(parsed);
    internalize_json_property(ctx, &holder, "", &reviver)
}

fn double_numbers_revive(value: Value) -> Value {
    match &value {
        Value::I32(n) => Value::I32(n * 2),
        Value::I64(n) => Value::I64(n * 2),
        Value::F64(n) => Value::F64(n * 2.0),
        Value::Object(obj) => {
            let guard = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = guard.kind {
                let elems2: Vec<Value> = elems
                    .iter()
                    .map(|e| double_numbers_revive(e.clone()))
                    .collect();
                drop(guard);
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems2)))
            } else {
                let keys: Vec<String> = guard.properties.keys().cloned().collect();
                let vals: Vec<(String, Value)> = keys
                    .iter()
                    .map(|k| {
                        (
                            k.clone(),
                            double_numbers_revive(guard.properties[k].clone()),
                        )
                    })
                    .collect();
                drop(guard);
                let mut new_obj = Object::new();
                for (k, v) in vals {
                    new_obj.properties.insert(k, v);
                }
                Value::Object(vybe_runtime::heap::alloc(new_obj))
            }
        }
        _ => value,
    }
}

fn internalize_json_property(
    ctx: &mut HostContext,
    holder: &Value,
    key: &str,
    reviver: &Value,
) -> Value {
    let mut value = get_holder_property(holder, key).unwrap_or(Value::Undefined);

    if let Value::Object(obj) = value.clone() {
        let is_array = {
            let guard = obj.lock().unwrap();
            matches!(guard.kind, ObjectKind::Array(_))
        };
        if is_array {
            let len = {
                let guard = obj.lock().unwrap();
                match &guard.kind {
                    ObjectKind::Array(elems) => elems.len(),
                    _ => 0,
                }
            };
            for index in 0..len {
                crate::keys::with_index_key(index, |idx_key| {
                    let revived = internalize_json_property(
                        ctx,
                        &Value::Object(obj.clone()),
                        idx_key,
                        reviver,
                    );
                    if matches!(revived, Value::Undefined) {
                        delete_holder_property(&Value::Object(obj.clone()), idx_key);
                    } else {
                        set_holder_property(&Value::Object(obj.clone()), idx_key, revived);
                    }
                });
            }
        } else {
            let keys = {
                let guard = obj.lock().unwrap();
                ordinary_ordered_keys(&guard)
                    .into_iter()
                    .filter(|name| is_serializable_object_key(&guard, name))
                    .collect::<Vec<_>>()
            };
            for child_key in keys {
                let revived = internalize_json_property(
                    ctx,
                    &Value::Object(obj.clone()),
                    &child_key,
                    reviver,
                );
                if matches!(revived, Value::Undefined) {
                    delete_holder_property(&Value::Object(obj.clone()), &child_key);
                } else {
                    set_holder_property(&Value::Object(obj.clone()), &child_key, revived);
                }
            }
        }

        value = Value::Object(obj);
    }

    crate::function::invoke_with_explicit_this(
        ctx,
        reviver,
        holder.clone(),
        &[crate::keys::string_value(key), value],
    )
}

fn get_holder_property(holder: &Value, key: &str) -> Option<Value> {
    let Value::Object(obj) = holder else {
        return None;
    };
    let guard = obj.lock().unwrap();
    if let Some(index) = json_array_index(key) {
        if let ObjectKind::Array(ref elems) = guard.kind {
            return elems.get(index as usize).cloned();
        }
    }
    guard.properties.get(key).cloned()
}

fn set_holder_property(holder: &Value, key: &str, value: Value) {
    let Value::Object(obj) = holder else {
        return;
    };
    let mut guard = obj.lock().unwrap();
    if let Some(index) = json_array_index(key) {
        if let ObjectKind::Array(ref mut elems) = guard.kind {
            if let Some(slot) = elems.get_mut(index as usize) {
                *slot = value;
                clear_array_hole(&mut guard, index as i32);
                return;
            }
        }
    }
    if let Some(slot) = guard.properties.get_mut(key) {
        *slot = value;
    } else {
        guard.properties.insert(key.to_string(), value);
    }
}

fn delete_holder_property(holder: &Value, key: &str) {
    let Value::Object(obj) = holder else {
        return;
    };
    let mut guard = obj.lock().unwrap();
    if let Some(index) = json_array_index(key) {
        if let ObjectKind::Array(ref mut elems) = guard.kind {
            if let Some(slot) = elems.get_mut(index as usize) {
                *slot = Value::Undefined;
                mark_array_hole(&mut guard, index as i32);
                return;
            }
        }
    }
    guard.properties.shift_remove(key);
}

fn mark_array_hole(object: &mut Object, index: i32) {
    let holes = match object.properties.get("__holes") {
        Some(Value::Object(existing)) => existing.clone(),
        _ => {
            let created = vybe_runtime::heap::alloc(Object::new_array(Vec::new()));
            object
                .properties
                .insert("__holes".into(), Value::Object(created.clone()));
            created
        }
    };

    let mut holes_guard = holes.lock().unwrap();
    let ObjectKind::Array(ref mut elems) = holes_guard.kind else {
        return;
    };
    if !elems
        .iter()
        .any(|value| matches!(value, Value::I32(existing) if *existing == index))
    {
        elems.push(Value::I32(index));
    }
}

fn clear_array_hole(object: &mut Object, index: i32) {
    let Some(Value::Object(holes)) = object.properties.get("__holes") else {
        return;
    };
    let mut holes_guard = holes.lock().unwrap();
    let ObjectKind::Array(ref mut elems) = holes_guard.kind else {
        return;
    };
    elems.retain(|value| !matches!(value, Value::I32(existing) if *existing == index));
}

fn parse_json(text: &str) -> Option<Value> {
    let mut p = Parser {
        src: text.as_bytes(),
        pos: 0,
    };
    p.skip_whitespace();
    let v = p.parse_value()?;
    p.skip_whitespace();
    if p.pos != p.src.len() {
        return None;
    }
    Some(v)
}

impl<'a> Parser<'a> {
    fn skip_whitespace(&mut self) {
        while self.pos < self.src.len() {
            match self.src[self.pos] {
                b' ' | b'\t' | b'\n' | b'\r' => self.pos += 1,
                _ => break,
            }
        }
    }

    fn peek(&self) -> Option<u8> {
        self.src.get(self.pos).copied()
    }

    fn parse_value(&mut self) -> Option<Value> {
        self.skip_whitespace();
        match self.peek()? {
            b'{' => self.parse_object(),
            b'[' => self.parse_array(),
            b'"' => self.parse_string().map(|text| match text {
                Cow::Borrowed(text) => crate::keys::string_value(text),
                Cow::Owned(text) => owned_string_value(text),
            }),
            b't' | b'f' => self.parse_bool(),
            b'n' => self.parse_null(),
            b'-' | b'0'..=b'9' => self.parse_number(),
            _ => None,
        }
    }

    fn parse_null(&mut self) -> Option<Value> {
        if self.src[self.pos..].starts_with(b"null") {
            self.pos += 4;
            Some(Value::Null)
        } else {
            None
        }
    }

    fn parse_bool(&mut self) -> Option<Value> {
        if self.src[self.pos..].starts_with(b"true") {
            self.pos += 4;
            Some(Value::Bool(true))
        } else if self.src[self.pos..].starts_with(b"false") {
            self.pos += 5;
            Some(Value::Bool(false))
        } else {
            None
        }
    }

    fn parse_number(&mut self) -> Option<Value> {
        let start = self.pos;
        if self.src.get(self.pos) == Some(&b'-') {
            self.pos += 1;
        }
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if !(c.is_ascii_digit()
                || c == b'.'
                || c == b'e'
                || c == b'E'
                || c == b'+'
                || c == b'-')
            {
                break;
            }
            self.pos += 1;
        }
        let s = std::str::from_utf8(&self.src[start..self.pos]).ok()?;
        // Try i64 first for exact integer preservation, then f64.
        if !s.contains('.') && !s.contains('e') && !s.contains('E') {
            if s == "-0" {
                return Some(Value::F64(-0.0));
            }
            if let Ok(n) = s.parse::<i64>() {
                // Fit into i32 if possible — matches v8's tagging
                // preference for small integers.
                if n >= i32::MIN as i64 && n <= i32::MAX as i64 {
                    return Some(Value::I32(n as i32));
                }
                return Some(Value::I64(n));
            }
        }
        s.parse::<f64>().ok().map(Value::F64)
    }

    /// The four hex digits of a `\uXXXX` escape starting at `at`.
    fn hex4_at(&self, at: usize) -> Option<u32> {
        let digits = self.src.get(at..at + 4)?;
        u32::from_str_radix(std::str::from_utf8(digits).ok()?, 16).ok()
    }

    fn parse_string(&mut self) -> Option<Cow<'a, str>> {
        if self.peek()? != b'"' {
            return None;
        }
        self.pos += 1;
        let start = self.pos;
        while self.pos < self.src.len() {
            match self.src[self.pos] {
                b'"' => {
                    let text = std::str::from_utf8(&self.src[start..self.pos]).ok()?;
                    self.pos += 1;
                    return Some(Cow::Borrowed(text));
                }
                b'\\' => break,
                byte if byte < 0x20 => return None,
                byte if byte < 0x80 => {
                    self.pos += 1;
                }
                _ => {
                    let remaining = &self.src[self.pos..];
                    let char_len = utf8_char_len(remaining[0]);
                    if char_len == 0 || self.pos + char_len > self.src.len() {
                        return None;
                    }
                    std::str::from_utf8(&remaining[..char_len]).ok()?;
                    self.pos += char_len;
                }
            }
        }
        let mut out = String::with_capacity(self.pos.saturating_sub(start));
        out.push_str(std::str::from_utf8(&self.src[start..self.pos]).ok()?);
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            match c {
                b'"' => {
                    self.pos += 1;
                    return Some(Cow::Owned(out));
                }
                b'\\' => {
                    self.pos += 1;
                    let esc = *self.src.get(self.pos)?;
                    self.pos += 1;
                    match esc {
                        b'"' => out.push('"'),
                        b'\\' => out.push('\\'),
                        b'/' => out.push('/'),
                        b'n' => out.push('\n'),
                        b't' => out.push('\t'),
                        b'r' => out.push('\r'),
                        b'b' => out.push('\x08'),
                        b'f' => out.push('\x0C'),
                        b'u' => {
                            let code = self.hex4_at(self.pos)?;
                            self.pos += 4;
                            // §25.5.1: a `\u` escape is a UTF-16 CODE UNIT, so a
                            // high surrogate combines with the `\u` escape that
                            // follows into one supplementary code point.
                            // `char::from_u32` rejects either half on its own —
                            // which silently dropped both, parsing
                            // `"😀"` to the empty string.
                            let code = match self.hex4_at(self.pos + 2) {
                                Some(low)
                                    if (0xD800..0xDC00).contains(&code)
                                        && (0xDC00..=0xDFFF).contains(&low)
                                        && self.src.get(self.pos) == Some(&b'\\')
                                        && self.src.get(self.pos + 1) == Some(&b'u') =>
                                {
                                    self.pos += 6;
                                    0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00)
                                }
                                _ => code,
                            };
                            // An unpaired surrogate is a code unit no `char` can
                            // hold; `wasm:js-string.fromCharCode` already renders
                            // one as U+FFFD via `from_utf16_lossy`, so match that
                            // rather than dropping the escape entirely.
                            out.push(char::from_u32(code).unwrap_or('\u{FFFD}'));
                        }
                        _ => return None,
                    }
                }
                _ => {
                    // Copy UTF-8 bytes up to the next special char.
                    // Cheapest: push one byte if ASCII, else advance
                    // and copy the full char via from_utf8.
                    let remaining = &self.src[self.pos..];
                    // Find the end of this char's byte sequence.
                    let char_len = utf8_char_len(remaining[0]);
                    if char_len == 0 || self.pos + char_len > self.src.len() {
                        return None;
                    }
                    let chunk = std::str::from_utf8(&remaining[..char_len]).ok()?;
                    out.push_str(chunk);
                    self.pos += char_len;
                }
            }
        }
        None // unterminated string
    }

    fn parse_array(&mut self) -> Option<Value> {
        if self.peek()? != b'[' {
            return None;
        }
        self.pos += 1;
        let mut elems: Vec<Value> = Vec::with_capacity(self.collection_capacity_hint());
        self.skip_whitespace();
        if self.peek() == Some(b']') {
            self.pos += 1;
            return Some(Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                elems,
            ))));
        }
        loop {
            elems.push(self.parse_value()?);
            self.skip_whitespace();
            match self.peek()? {
                b',' => {
                    self.pos += 1;
                }
                b']' => {
                    self.pos += 1;
                    break;
                }
                _ => return None,
            }
        }
        Some(Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            elems,
        ))))
    }

    fn parse_object(&mut self) -> Option<Value> {
        if self.peek()? != b'{' {
            return None;
        }
        self.pos += 1;
        let capacity_hint = self.collection_capacity_hint();
        let mut obj = Object::new();
        obj.properties.reserve(capacity_hint);
        let mut tracked_keys: Vec<Value> = Vec::with_capacity(capacity_hint);
        self.skip_whitespace();
        if self.peek() == Some(b'}') {
            self.pos += 1;
            return Some(Value::Object(vybe_runtime::heap::alloc(obj)));
        }
        loop {
            self.skip_whitespace();
            let key = self.parse_string()?.into_owned();
            self.skip_whitespace();
            if self.peek() != Some(b':') {
                return None;
            }
            self.pos += 1;
            let val = self.parse_value()?;
            if !obj.properties.contains_key(&key) {
                tracked_keys.push(crate::keys::string_value(&key));
            }
            obj.properties.insert(key, val);
            self.skip_whitespace();
            match self.peek()? {
                b',' => {
                    self.pos += 1;
                }
                b'}' => {
                    self.pos += 1;
                    break;
                }
                _ => return None,
            }
        }
        obj.properties.insert(
            "__keys".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(tracked_keys))),
        );
        Some(Value::Object(vybe_runtime::heap::alloc(obj)))
    }

    #[inline]
    fn collection_capacity_hint(&self) -> usize {
        self.src
            .len()
            .saturating_sub(self.pos)
            .saturating_div(4)
            .min(JSON_PARSE_PREALLOC_LIMIT)
    }
}

/// UTF-8 leading-byte → sequence length. Returns 0 on invalid.
fn utf8_char_len(b: u8) -> usize {
    if b & 0x80 == 0 {
        1
    } else if b & 0xE0 == 0xC0 {
        2
    } else if b & 0xF0 == 0xE0 {
        3
    } else if b & 0xF8 == 0xF0 {
        4
    } else {
        0
    }
}

#[cfg(test)]
mod borrowed_json_strings_tests {
    use super::*;

    fn decode(input: &str) -> Cow<'_, str> {
        Parser { src: input.as_bytes(), pos: 0 }.parse_string().unwrap()
    }

    #[test]
    fn unescaped_strings_borrow_and_escaped_strings_own() {
        for (input, expected) in [
            ("\"\"", ""),
            ("\"plain ASCII\"", "plain ASCII"),
            ("\"\u{e9}\u{1f600}\"", "\u{e9}\u{1f600}"),
        ] {
            let decoded = decode(input);
            assert!(matches!(decoded, Cow::Borrowed(_)));
            assert_eq!(decoded.as_ref(), expected);
        }
        for (input, expected) in [
            (r#""line\nend""#, "line\nend"),
            (r#""\"\\\/\b\f\r\t""#, "\"\\/\x08\x0c\r\t"),
            (r#""\u0061\uD83D\uDE00""#, "a\u{1f600}"),
        ] {
            let decoded = decode(input);
            assert!(matches!(decoded, Cow::Owned(_)));
            assert_eq!(decoded.as_ref(), expected);
        }
    }

    #[test]
    fn runtime_strings_outlive_parser_input() {
        for (input, expected) in [
            ("\"plain value\"", "plain value"),
            (r#""escaped\nvalue""#, "escaped\nvalue"),
        ] {
            let value = {
                let source = input.to_owned();
                parse_json(&source).unwrap()
            };
            let Value::String(text) = value else { panic!("expected runtime string") };
            assert_eq!(text.as_ref(), expected);
        }
    }

    #[test]
    fn decoded_duplicate_keys_keep_last_value_and_original_order() {
        let value = parse_json(r#"{"escaped\u006bey":1,"escapedkey":2,"plain":3}"#).unwrap();
        let Value::Object(object) = value else { panic!("expected object") };
        let object = object.lock().unwrap();
        assert_eq!(object.properties.get("escapedkey").unwrap().as_i32(), 2);
        let Value::Object(keys) = object.properties.get("__keys").unwrap() else {
            panic!("expected tracked keys")
        };
        let keys = keys.lock().unwrap();
        let ObjectKind::Array(keys) = &keys.kind else { panic!("expected key array") };
        assert_eq!(keys.len(), 2);
        assert!(matches!(&keys[0], Value::String(key) if key.as_ref() == "escapedkey"));
        assert!(matches!(&keys[1], Value::String(key) if key.as_ref() == "plain"));
    }

    #[test]
    fn invalid_fast_path_strings_still_fail() {
        for input in ["\"unterminated", "\"bad\ntext\"", r#""\x20""#, r#""\uZZZZ""#] {
            assert!(parse_json(input).is_none(), "{input:?}");
        }
    }

    fn stringify_for_test(value: Value, space: i32) -> String {
        let mut vm = VM::new();
        register(&mut vm);
        let mut callable = Object::new();
        callable.kind = ObjectKind::HostFunction(
            vm.resolve_host_function_index("ecma:json", "stringify").unwrap(),
        );
        let result = vm.invoke(
            &Value::Object(vybe_runtime::heap::alloc(callable)),
            &[value, Value::Undefined, Value::I32(space)],
        ).unwrap();
        let Value::String(text) = result else { panic!("expected serialized string") };
        text.to_string()
    }

    #[test]
    fn streamed_array_preserves_primitive_text_and_null_omission() {
        let values = vec![
            Value::Bool(true), Value::Bool(false), Value::Null, Value::Undefined,
            Value::I32(7), crate::keys::string_value("line\n\"end"),
        ];
        assert_eq!(stringify_for_test(
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))), 0,
        ), "[true,false,null,null,7,\"line\\n\\\"end\"]");
    }

    #[test]
    fn streamed_object_omission_preserves_separators() {
        let mut object = Object::new();
        object.properties.insert("first".into(), Value::Bool(true));
        object.properties.insert("omitted".into(), Value::Undefined);
        object.properties.insert("last".into(), crate::keys::string_value("text"));
        assert_eq!(stringify_for_test(Value::Object(vybe_runtime::heap::alloc(object)), 0),
            "{\"first\":true,\"last\":\"text\"}");
    }

    #[test]
    fn pretty_indentation_uses_one_gap_per_level() {
        assert_eq!(stringify_for_test(Value::Object(vybe_runtime::heap::alloc(
            Object::new_array(vec![Value::I32(1), Value::Bool(true)]),
        )), 2), "[\n  1,\n  true\n]");
        let mut inner = Object::new();
        inner.properties.insert("value".into(), Value::I32(7));
        let mut outer = Object::new();
        outer.properties.insert("first".into(), Value::Bool(true));
        outer.properties.insert("omitted".into(), Value::Undefined);
        outer.properties.insert("last".into(), Value::Object(vybe_runtime::heap::alloc(inner)));
        assert_eq!(stringify_for_test(Value::Object(vybe_runtime::heap::alloc(outer)), 2),
            "{\n  \"first\": true,\n  \"last\": {\n    \"value\": 7\n  }\n}");
    }

    #[test]
    #[ignore = "matched primitive writer microbenchmark, not complete JSON.stringify timing"]
    fn native_streamed_primitive_microbenchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        let mut vm = VM::new();
        vm.register_free_fn("json-benchmark", "run", Box::new(|ctx, _| {
            let values = [Value::Bool(true), Value::Bool(false), Value::Null,
                Value::I32(7), crate::keys::string_value("line\n\"end")];
            let holder = make_root_holder(Value::Null);
            let mut run = |streamed: bool| {
                let mut state = StringifyState::new(None, None);
                let mut out = String::with_capacity(128);
                let start = Instant::now();
                for _ in 0..20_000 {
                    out.clear();
                    out.push('[');
                    for (index, value) in values.iter().enumerate() {
                        crate::keys::with_index_key(index, |key| {
                            if streamed {
                                assert!(serialize_property_into(ctx, &holder, key,
                                    value.clone(), &mut state, true, &mut out,
                                    |out| { if index > 0 { out.push(','); } }));
                            } else {
                                let text = serialize_property(ctx, &holder, key,
                                    value.clone(), &mut state, true).unwrap();
                                if index > 0 { out.push(','); }
                                out.push_str(&text);
                            }
                        });
                    }
                    out.push(']');
                    black_box(&out);
                }
                assert_eq!(out, "[true,false,null,7,\"line\\n\\\"end\"]");
                start.elapsed()
            };
            for pair in 0..4 {
                let (allocated, streamed) = if pair % 2 == 0 {
                    let allocated = run(false);
                    (allocated, run(true))
                } else {
                    let streamed = run(true);
                    (run(false), streamed)
                };
                eprintln!("JSON primitive writer pair {pair}: allocated={allocated:?} streamed={streamed:?}");
            }
            Value::Null
        }));
        let mut callable = Object::new();
        callable.kind = ObjectKind::HostFunction(
            vm.resolve_host_function_index("json-benchmark", "run").unwrap(),
        );
        vm.invoke(&Value::Object(vybe_runtime::heap::alloc(callable)), &[]).unwrap();
    }

    #[test]
    #[ignore = "matched decoder/value-construction microbenchmark, not complete JSON.parse timing"]
    fn native_borrowed_string_value_microbenchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        fn run(input: &str, copy: bool, iterations: usize) -> std::time::Duration {
            let start = Instant::now();
            for _ in 0..iterations {
                let text = decode(black_box(input));
                let value = if copy {
                    crate::keys::string_value(&text.into_owned())
                } else {
                    match text {
                        Cow::Borrowed(text) => crate::keys::string_value(text),
                        Cow::Owned(text) => owned_string_value(text),
                    }
                };
                black_box(value);
            }
            start.elapsed()
        }
        let input = format!("\"{}\"", "ascii payload ".repeat(16));
        for pair in 0..4 {
            let (copied, borrowed) = if pair % 2 == 0 {
                let copied = run(&input, true, 100_000);
                (copied, run(&input, false, 100_000))
            } else {
                let borrowed = run(&input, false, 100_000);
                (run(&input, true, 100_000), borrowed)
            };
            eprintln!("JSON string decode/value construction pair {pair}: copied={copied:?} borrowed={borrowed:?}");
        }
    }
}

#[cfg(test)]
mod full_operation_performance_tests {
    use super::*;
    use std::time::Instant;

    fn host_function(vm: &VM, name: &str) -> Value {
        let idx = *vm.host_registry
            .get(&("ecma:json".to_owned(), name.to_owned()))
            .expect("JSON host function registered");
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(idx);
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    fn run_case(vm: &mut VM, callable: &Value, input: &Value, stringify: bool) -> Value {
        let args = if stringify {
            vec![input.clone(), Value::Undefined, Value::Undefined]
        } else {
            vec![input.clone(), Value::Undefined]
        };
        vm.invoke(callable, &args).expect("JSON operation succeeds")
    }

    #[test]
    #[ignore = "full registered-adapter benchmark; run explicitly with --nocapture"]
    fn registered_json_full_operations() {
        let mut vm = VM::new();
        register(&mut vm);
        let stringify = host_function(&vm, "stringify");
        let parse = host_function(&vm, "parse");
        let values = (0..1000).map(|i| match i % 3 {
            0 => crate::keys::string_value("line\n\"value"),
            1 => Value::Bool(true),
            _ => Value::I32(i),
        }).collect();
        let array = Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)));
        let mut object = Object::new();
        for i in 0..500 {
            object.properties.insert(format!("field{i}"), if i % 2 == 0 {
                crate::keys::string_value("value")
            } else {
                Value::I32(i)
            });
        }
        let object = Value::Object(vybe_runtime::heap::alloc(object));
        let array_text = run_case(&mut vm, &stringify, &array, true);
        let object_text = run_case(&mut vm, &stringify, &object, true);
        let cases = [
            ("stringify-array", &stringify, &array, true, &array_text),
            ("stringify-object", &stringify, &object, true, &object_text),
            ("parse-array", &parse, &array_text, false, &array_text),
            ("parse-object", &parse, &object_text, false, &object_text),
        ];
        for (name, callable, input, is_stringify, expected_text) in cases {
            for _ in 0..10 {
                std::hint::black_box(run_case(&mut vm, callable, input, is_stringify));
            }
            for sample in 0..4 {
                let start = Instant::now();
                let mut last = Value::Undefined;
                for _ in 0..200 {
                    last = run_case(&mut vm, callable, input, is_stringify);
                    std::hint::black_box(&last);
                }
                let elapsed = start.elapsed();
                let text = if is_stringify { last } else {
                    run_case(&mut vm, &stringify, &last, true)
                };
                assert_eq!(&text, expected_text, "{name} result differs");
                eprintln!("{name} sample={sample} operations=200 elapsed={elapsed:?}");
            }
        }
    }
}

#[cfg(test)]
mod borrowed_primitive_array_tests {
    use super::*;

    #[test]
    #[ignore = "paired primitive-array serialization kernels through registered callbacks"]
    fn borrowed_array_serialization_benchmark() {
        let mut vm = VM::new();
        for borrowed in [false, true] {
            vm.register_free_fn("ecma:array-serialization-check", if borrowed { "borrowed" } else { "previous" },
                Box::new(move |ctx, args| {
                    let args = ctx.user_args(args, 0);
                    let Value::Object(object) = &args[0] else { unreachable!() };
                    let object_guard = object.lock().unwrap();
                    let ObjectKind::Array(elems) = &object_guard.kind else { unreachable!() };
                    let mut state = StringifyState::new(None, None);
                    let output = if borrowed {
                        serialize_compact_primitive_array(&object_guard, elems, &state).unwrap()
                    } else {
                        let elems = elems.clone();
                        drop(object_guard);
                        serialize_array(ctx, object, elems, &mut state).unwrap()
                    };
                    crate::keys::string_value(&output)
                }));
        }
        let functions = ["previous", "borrowed"].map(|name| {
            let mut object = Object::new();
            object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:array-serialization-check", name).unwrap());
            Value::Object(vybe_runtime::heap::alloc(object))
        });
        for strings in [false, true] {
            let elems = (0..1024).map(|index| if strings {
                crate::keys::string_value("value\"escaped")
            } else { Value::I32(index) }).collect();
            let input = [Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))];
            let old = vm.invoke(&functions[0], &input).unwrap();
            let new = vm.invoke(&functions[1], &input).unwrap();
            let (Value::String(old), Value::String(new)) = (old, new) else { unreachable!() };
            assert_eq!(old, new);
            let mut times = [std::time::Duration::ZERO; 2];
            for round in 0..6 {
                for variant in if round % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let start = std::time::Instant::now();
                    for _ in 0..200 {
                        std::hint::black_box(vm.invoke(&functions[variant], &input).unwrap());
                    }
                    times[variant] += start.elapsed();
                }
            }
            eprintln!("primitive-array strings={strings} elements=1024 calls=1200 previous={:?} borrowed={:?} ratio={:.3}",
                times[0], times[1], times[0].as_secs_f64() / times[1].as_secs_f64());
        }
    }

    #[test]
    fn compact_primitives_preserve_nulls_numbers_and_escaping() {
        let object = Object::new_array(vec![Value::Undefined, Value::Null,
            Value::Bool(true), Value::I32(-7), Value::I64(42),
            Value::F64(-0.0), Value::F64(f64::NAN), Value::F32(f32::INFINITY),
            crate::keys::string_value("quote\"\n")]);
        let ObjectKind::Array(elems) = &object.kind else { unreachable!() };
        assert_eq!(serialize_compact_primitive_array(&object, elems, &StringifyState::new(None, None)),
            Some("[null,null,true,-7,42,0,null,null,\"quote\\\"\\n\"]".into()));
    }

    #[test]
    fn observable_or_native_cases_keep_the_existing_path() {
        let mut object = Object::new_array(vec![Value::I32(1)]);
        let state = StringifyState::new(None, None);
        object.properties.insert("__holes".into(), Value::Bool(true));
        let ObjectKind::Array(elems) = &object.kind else { unreachable!() };
        assert!(serialize_compact_primitive_array(&object, elems, &state).is_none());
        object.properties.clear();
        object.type_id = 1;
        let ObjectKind::Array(elems) = &object.kind else { unreachable!() };
        assert!(serialize_compact_primitive_array(&object, elems, &state).is_none());
        object.type_id = 0;
        let mut state = state;
        state.gap = " ".into();
        let ObjectKind::Array(elems) = &object.kind else { unreachable!() };
        assert!(serialize_compact_primitive_array(&object, elems, &state).is_none());
        state.gap.clear();
        state.replacer_fn = Some(Value::Undefined);
        assert!(serialize_compact_primitive_array(&object, elems, &state).is_none());
        state.replacer_fn = None;
        let nested = [Value::Object(vybe_runtime::heap::alloc(Object::new()))];
        assert!(serialize_compact_primitive_array(&object, &nested, &state).is_none());
    }
}

#[cfg(test)]
mod streamed_number_tests {
    use super::*;

fn previous_json_number_string(n: f64) -> String {
    if n.is_nan() || n.is_infinite() {
        return "null".to_string();
    }
    if n == 0.0 {
        return "0".to_string();
    }
    if n.fract() == 0.0 && n.abs() < 1e16 {
        return (n as i64).to_string();
    }
    n.to_string()
}

    #[test]
    fn streaming_preserves_existing_number_bytes_at_boundaries_and_across_bits() {
        let fixed = [f64::NAN, f64::INFINITY, f64::NEG_INFINITY, 0.0, -0.0,
            0.125, -0.125, 1e-7, 1e-6, 1e15, 1e16, 1e20, 1e21,
            f64::MAX, f64::MIN_POSITIVE, f64::from_bits(1)];
        let mut state = 0x9e3779b97f4a7c15u64;
        for value in fixed.into_iter().chain((0..8192).map(|_| {
            state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
            f64::from_bits(state)
        })) {
            let expected = previous_json_number_string(value);
            let mut output = String::from("prefix:");
            write_json_number(value, &mut output);
            assert_eq!(&output[7..], expected, "bits={:016x}", value.to_bits());
            assert_eq!(json_number_string(value), expected);
        }
    }

    #[test]
    #[ignore = "paired registered numeric-array serialization kernels; excludes generic JSON callback plumbing"]
    fn registered_numeric_array_stream_benchmark() {
        let mut vm = VM::new();
        for stream in [false, true] {
            vm.register_free_fn("ecma:number-stream-check", if stream { "stream" } else { "previous" },
                Box::new(move |_, args| {
                    let Value::Object(input) = &args[0] else { unreachable!() };
                    let guard = input.lock().unwrap();
                    let ObjectKind::Array(values) = &guard.kind else { unreachable!() };
                    let mut output = String::with_capacity(values.len().saturating_mul(12));
                    output.push('[');
                    for (index, value) in values.iter().enumerate() {
                        if index != 0 { output.push(','); }
                        let Value::F64(number) = value else { unreachable!() };
                        if stream {
                            write_json_number(*number, &mut output);
                        } else {
                            output.push_str(&previous_json_number_string(*number));
                        }
                    }
                    output.push(']');
                    crate::keys::owned_string_value(output)
                }));
        }
        let functions = ["previous", "stream"].map(|name| {
            let mut object = Object::new();
            object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:number-stream-check", name).unwrap());
            Value::Object(vybe_runtime::heap::alloc(object))
        });
        for fractional in [false, true] {
            let values = (0..1024).map(|i| Value::F64(i as f64 + if fractional { 0.125 } else { 0.0 })).collect();
            let args = [Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))];
            let old = vm.invoke(&functions[0], &args).unwrap();
            let new = vm.invoke(&functions[1], &args).unwrap();
            let (Value::String(old), Value::String(new)) = (old, new) else { unreachable!() };
            assert_eq!(old, new);
            let mut times = [std::time::Duration::ZERO; 2];
            for round in 0..8 {
                for variant in if round % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let started = std::time::Instant::now();
                    for _ in 0..200 {
                        std::hint::black_box(vm.invoke(&functions[variant], &args).unwrap());
                    }
                    times[variant] += started.elapsed();
                }
            }
            eprintln!("json-number-stream fractional={fractional} elements=1024 calls=1600 previous={:?} stream={:?} ratio={:.3}",
                times[0], times[1], times[0].as_secs_f64() / times[1].as_secs_f64());
        }
    }
}

