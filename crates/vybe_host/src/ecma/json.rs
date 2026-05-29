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
//!   - Circular references are detected via a visited-set. MVP keeps the
//!     existing non-throwing behavior and serializes the cycle as null.
//!
//! See `JS_BUILTIN_CONVENTIONS.md`.

use std::collections::HashSet;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, TypedElemKind, Value};
use vybe_bytecode::{HostContext, VM};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("ecma:json", "stringify",
        Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let mut state = StringifyState::new(args.get(1), args.get(2));
            let root_holder = make_root_holder(value.clone());
            match serialize_property(ctx, &root_holder, "", value, &mut state, false) {
                Some(text) => Value::String(Arc::from(text.as_str())),
                None => Value::Undefined,
            }
        }));

    vm.register_host_fn("ecma:json", "parse",
        Box::new(|ctx, args| {
            let text: String = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Undefined,
            };
            let parsed = parse_json(&text).unwrap_or(Value::Null);
            match args.get(1).cloned() {
                Some(reviver) if is_callable(&reviver) => apply_reviver(ctx, parsed, reviver),
                _ => parsed,
            }
        }));

    // JSON.stringify(value, replacer, space) — replacer as Array filters keys.
    vm.register_host_fn("ecma:json", "stringifyWithReplacer",
        Box::new(|ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            let replacer = args.get(1).cloned();
            let space = args.get(2).cloned();
            let mut state = StringifyState::new(replacer.as_ref(), space.as_ref());
            let root_holder = make_root_holder(value.clone());
            match serialize_property(ctx, &root_holder, "", value, &mut state, false) {
                Some(text) => Value::String(Arc::from(text.as_str())),
                None => Value::Undefined,
            }
        }));

    // JSON.parse(text, reviver) — reviver transforms each member.
    vm.register_host_fn("ecma:json", "parseWithReviver",
        Box::new(|ctx, args| {
            let text: String = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Undefined,
            };
            let parsed = parse_json(&text).unwrap_or(Value::Null);
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
        }));

    // JSON.rawJSON(text) — ES2025: wraps a raw JSON text in an opaque object.
    vm.register_host_fn("ecma:json", "rawJSON",
        Box::new(|_ctx, args| {
            let text = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Undefined,
            };
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), Value::String(Arc::from("RawJSON")));
            obj.properties.insert("rawJSON".into(), Value::String(Arc::from(text.as_str())));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // JSON.isRawJSON(value) — ES2025: true iff value was created by JSON.rawJSON().
    vm.register_host_fn("ecma:json", "isRawJSON",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if o.properties.get("__type").map(|v| format!("{}", v)).as_deref() == Some("RawJSON") {
                    return Value::Bool(true);
                }
            }
            Value::Bool(false)
        }));
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
        value = crate::ecma::function::invoke_with_explicit_this(
            ctx,
            replacer,
            holder.clone(),
            &[Value::String(Arc::from(key)), value],
        );
    }

    unbox_json_wrapper(value)
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
        Value::Null => Some("null".to_string()),
        Value::Bool(b) => Some(if *b { "true".into() } else { "false".into() }),
        Value::I32(n) => Some(n.to_string()),
        Value::I64(n) => Some(n.to_string()),
        Value::F64(n) => Some(json_number_string(*n)),
        Value::String(s) => Some(quote_string(s)),
        Value::V128(_) | Value::WeakRef(_) => Some("null".to_string()),
        // ECMA-262 §25.5.2: BigInt values must throw TypeError.
        Value::BigInt(_) => {
            let err = crate::ecma::error::new_error("TypeError", "Do not know how to serialize a BigInt");
            ctx.throw_value(err);
            return None;
        }
        Value::Object(obj) => serialize_object(ctx, obj, state),
    }
}

fn serialize_object(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    state: &mut StringifyState,
) -> Option<String> {
    let id = Arc::as_ptr(obj) as usize;
    if !state.visited.insert(id) {
        return Some("null".to_string());
    }

    let result = {
        let guard = obj.lock().unwrap();
        // RawJSON objects serialize their embedded literal directly.
        if guard.properties.get("__type").map(|v| format!("{}", v)).as_deref() == Some("RawJSON") {
            if let Some(Value::String(raw)) = guard.properties.get("rawJSON") {
                return Some(raw.to_string());
            }
        }
        match &guard.kind {
            ObjectKind::Array(elems) => serialize_array(ctx, obj, elems.clone(), state),
            ObjectKind::TypedArray(ta) => Some(stringify_typed_array(ta)),
            ObjectKind::Map(_) | ObjectKind::Set(_) | ObjectKind::ArrayBuffer(_) => {
                Some("{}".to_string())
            }
            ObjectKind::Function(_) | ObjectKind::HostFunction(_) | ObjectKind::Continuation(_) => None,
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

fn serialize_array(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    elems: Vec<Value>,
    state: &mut StringifyState,
) -> Option<String> {
    let holder = Value::Object(obj.clone());
    let stepback = state.indent.clone();
    state.indent.push_str(&state.gap);

    let mut parts = Vec::with_capacity(elems.len());
    for (index, value) in elems.into_iter().enumerate() {
        let key = index.to_string();
        parts.push(
            serialize_property(ctx, &holder, &key, value, state, true)
                .unwrap_or_else(|| "null".to_string()),
        );
    }

    state.indent = stepback.clone();
    if state.gap.is_empty() {
        return Some(format!("[{}]", parts.join(",")));
    }
    if parts.is_empty() {
        return Some("[]".to_string());
    }

    let body = parts
        .iter()
        .map(|part| format!("{}{}", state.indent.clone() + &state.gap, part))
        .collect::<Vec<_>>()
        .join(",\n");
    Some(format!("[\n{}\n{}]", body, stepback))
}

fn stringify_typed_array(ta: &vybe_bytecode::value::TypedArrayState) -> String {
    // Typed arrays stringify as the comma-joined element values
    // wrapped in an object shape — actually JSON.stringify on a typed
    // array produces a plain object with numeric-string keys. v8:
    //   JSON.stringify(new Int32Array([1,2,3])) === '{"0":1,"1":2,"2":3}'
    let buf = ta.buffer.lock().unwrap();
    let bpe = ta.elem.bytes_per_element();
    let available_elems = if ta.byte_offset >= buf.len() { 0 }
        else { (buf.len() - ta.byte_offset) / bpe };
    let length = ta.length.min(available_elems);
    let mut out = String::from("{");
    for i in 0..length {
        if i > 0 { out.push(','); }
        out.push_str(&format!("\"{}\":", i));
        let abs = ta.byte_offset + i * bpe;
        let val_str = match ta.elem {
            TypedElemKind::I8  => (buf[abs] as i8).to_string(),
            TypedElemKind::U8 | TypedElemKind::U8Clamped => buf[abs].to_string(),
            TypedElemKind::I16 => {
                let b = [buf[abs], buf[abs + 1]];
                i16::from_le_bytes(b).to_string()
            }
            TypedElemKind::U16 => {
                let b = [buf[abs], buf[abs + 1]];
                u16::from_le_bytes(b).to_string()
            }
            TypedElemKind::I32 => {
                let mut b = [0u8; 4]; b.copy_from_slice(&buf[abs..abs + 4]);
                i32::from_le_bytes(b).to_string()
            }
            TypedElemKind::U32 => {
                let mut b = [0u8; 4]; b.copy_from_slice(&buf[abs..abs + 4]);
                u32::from_le_bytes(b).to_string()
            }
            TypedElemKind::F32 => {
                let mut b = [0u8; 4]; b.copy_from_slice(&buf[abs..abs + 4]);
                let f = f32::from_le_bytes(b);
                if f.is_nan() || f.is_infinite() { "null".into() } else { f.to_string() }
            }
            TypedElemKind::F64 => {
                let mut b = [0u8; 8]; b.copy_from_slice(&buf[abs..abs + 8]);
                let f = f64::from_le_bytes(b);
                if f.is_nan() || f.is_infinite() { "null".into() } else { f.to_string() }
            }
            TypedElemKind::BigI64 => {
                let mut b = [0u8; 8]; b.copy_from_slice(&buf[abs..abs + 8]);
                i64::from_le_bytes(b).to_string()
            }
            TypedElemKind::BigU64 => {
                let mut b = [0u8; 8]; b.copy_from_slice(&buf[abs..abs + 8]);
                u64::from_le_bytes(b).to_string()
            }
        };
        out.push_str(&val_str);
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

    let mut parts = Vec::new();
    for key in keys {
        let value = {
            let guard = obj.lock().unwrap();
            guard.properties.get(key).cloned()
        };
        let Some(value) = value else {
            continue;
        };
        if let Some(serialized) = serialize_property(ctx, &holder, key, value, state, false) {
            let member = if state.gap.is_empty() {
                format!("{}:{}", quote_string(key), serialized)
            } else {
                format!("{}: {}", quote_string(key), serialized)
            };
            parts.push(member);
        }
    }

    state.indent = stepback.clone();
    if state.gap.is_empty() {
        return format!("{{{}}}", parts.join(","));
    }
    if parts.is_empty() {
        return "{}".to_string();
    }

    let body = parts
        .iter()
        .map(|part| format!("{}{}", state.indent.clone() + &state.gap, part))
        .collect::<Vec<_>>()
        .join(",\n");
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
            elems.iter()
                .filter_map(|elem| match elem {
                    Value::String(key) if o.properties.contains_key(key.as_ref()) => Some(key.to_string()),
                    _ => None,
                })
                .collect(),
        )
    });

    let live: Vec<String> = o.properties.keys().cloned().collect();
    match tracked {
        Some(mut keys) => {
            let mut seen: HashSet<String> = keys.iter().cloned().collect();
            for key in live {
                if seen.insert(key.clone()) {
                    keys.push(key);
                }
            }
            keys
        }
        None => live,
    }
}

fn object_serialization_keys(o: &Object, property_list: Option<&Vec<String>>) -> Vec<String> {
    let ordered = ordinary_ordered_keys(o);
    if let Some(list) = property_list {
        return list
            .iter()
            .filter(|key| is_serializable_object_key(o, key) && o.properties.contains_key(*key))
            .cloned()
            .collect();
    }

    let mut indices = Vec::new();
    let mut others = Vec::new();
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
            elems.iter().any(|value| matches!(value, Value::String(name) if name.as_ref() == key))
        }
        _ => false,
    }
}

fn json_array_index(key: &str) -> Option<u32> {
    if key.is_empty() || (key.len() > 1 && key.starts_with('0')) {
        return None;
    }
    let parsed = key.parse::<u32>().ok()?;
    if parsed == u32::MAX {
        return None;
    }
    if parsed.to_string() == key {
        Some(parsed)
    } else {
        None
    }
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

    let mut keys = Vec::new();
    let mut seen = HashSet::new();
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
    match unbox_json_wrapper(value.clone()) {
        Value::String(text) => Some(text.to_string()),
        Value::I32(n) => Some(n.to_string()),
        Value::I64(n) => Some(n.to_string()),
        Value::F64(n) if n.is_finite() => Some(json_number_string(n)),
        _ => None,
    }
}

fn json_number_string(n: f64) -> String {
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

fn is_callable(value: &Value) -> bool {
    let Value::Object(obj) = value else {
        return false;
    };
    let guard = obj.lock().unwrap();
    matches!(guard.kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_))
        || guard.properties.contains_key("__call__")
}

fn is_function_like_object(obj: &Arc<Mutex<Object>>) -> bool {
    let guard = obj.lock().unwrap();
    matches!(guard.kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_))
        || guard.properties.contains_key("__call__")
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
            return Some(crate::ecma::function::invoke_with_explicit_this(
                ctx,
                &method,
                value.clone(),
                &[Value::String(Arc::from(key))],
            ));
        }
    }

    let is_date = {
        let guard = obj.lock().unwrap();
        matches!(guard.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "Date")
    };
    if is_date {
        return crate::ecma::date::dispatch_date_method("toJSON", &[value.clone()]);
    }
    None
}

fn make_root_holder(value: Value) -> Value {
    let mut root = Object::new();
    root.properties.insert("".into(), value);
    Value::Object(Arc::new(Mutex::new(root)))
}

fn quote_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
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
                out.push_str(&format!("\\u{:04x}", c as u32));
            }
            c => out.push(c),
        }
    }
    out.push('"');
    out
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
                let elems2: Vec<Value> = elems.iter().map(|e| double_numbers_revive(e.clone())).collect();
                drop(guard);
                Value::Object(Arc::new(Mutex::new(Object::new_array(elems2))))
            } else {
                let keys: Vec<String> = guard.properties.keys().cloned().collect();
                let vals: Vec<(String, Value)> = keys.iter()
                    .map(|k| (k.clone(), double_numbers_revive(guard.properties[k].clone())))
                    .collect();
                drop(guard);
                let mut new_obj = Object::new();
                for (k, v) in vals { new_obj.properties.insert(k, v); }
                Value::Object(Arc::new(Mutex::new(new_obj)))
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
                let idx_key = index.to_string();
                let revived = internalize_json_property(
                    ctx,
                    &Value::Object(obj.clone()),
                    &idx_key,
                    reviver,
                );
                if matches!(revived, Value::Undefined) {
                    delete_holder_property(&Value::Object(obj.clone()), &idx_key);
                } else {
                    set_holder_property(&Value::Object(obj.clone()), &idx_key, revived);
                }
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

    crate::ecma::function::invoke_with_explicit_this(
        ctx,
        reviver,
        holder.clone(),
        &[Value::String(Arc::from(key)), value],
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
    guard.properties.insert(key.to_string(), value);
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
    guard.properties.remove(key);
}

fn mark_array_hole(object: &mut Object, index: i32) {
    let holes = match object.properties.get("__holes") {
        Some(Value::Object(existing)) => existing.clone(),
        _ => {
            let created = Arc::new(Mutex::new(Object::new_array(Vec::new())));
            object.properties.insert("__holes".into(), Value::Object(created.clone()));
            created
        }
    };

    let mut holes_guard = holes.lock().unwrap();
    let ObjectKind::Array(ref mut elems) = holes_guard.kind else {
        return;
    };
    if !elems.iter().any(|value| matches!(value, Value::I32(existing) if *existing == index)) {
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
    let mut p = Parser { src: text.as_bytes(), pos: 0 };
    p.skip_whitespace();
    let v = p.parse_value()?;
    p.skip_whitespace();
    if p.pos != p.src.len() {
        // Trailing content — per spec this is a SyntaxError; MVP returns null.
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
            b'"' => self.parse_string().map(|s| Value::String(Arc::from(s.as_str()))),
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
        if self.src.get(self.pos) == Some(&b'-') { self.pos += 1; }
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if !(c.is_ascii_digit() || c == b'.' || c == b'e' || c == b'E'
                || c == b'+' || c == b'-')
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

    fn parse_string(&mut self) -> Option<String> {
        if self.peek()? != b'"' { return None; }
        self.pos += 1;
        let mut out = String::new();
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            match c {
                b'"' => {
                    self.pos += 1;
                    return Some(out);
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
                            if self.pos + 4 > self.src.len() { return None; }
                            let hex = std::str::from_utf8(&self.src[self.pos..self.pos + 4]).ok()?;
                            let code = u32::from_str_radix(hex, 16).ok()?;
                            self.pos += 4;
                            if let Some(ch) = char::from_u32(code) {
                                out.push(ch);
                            }
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
        if self.peek()? != b'[' { return None; }
        self.pos += 1;
        let mut elems: Vec<Value> = Vec::new();
        self.skip_whitespace();
        if self.peek() == Some(b']') {
            self.pos += 1;
            return Some(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))));
        }
        loop {
            elems.push(self.parse_value()?);
            self.skip_whitespace();
            match self.peek()? {
                b',' => { self.pos += 1; }
                b']' => { self.pos += 1; break; }
                _ => return None,
            }
        }
        Some(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))
    }

    fn parse_object(&mut self) -> Option<Value> {
        if self.peek()? != b'{' { return None; }
        self.pos += 1;
        let mut obj = Object::new();
        let mut tracked_keys: Vec<Value> = Vec::new();
        let mut seen_keys: HashSet<String> = HashSet::new();
        self.skip_whitespace();
        if self.peek() == Some(b'}') {
            self.pos += 1;
            return Some(Value::Object(Arc::new(Mutex::new(obj))));
        }
        loop {
            self.skip_whitespace();
            let key = self.parse_string()?;
            self.skip_whitespace();
            if self.peek() != Some(b':') { return None; }
            self.pos += 1;
            let val = self.parse_value()?;
            if seen_keys.insert(key.clone()) {
                tracked_keys.push(Value::String(Arc::from(key.as_str())));
            }
            obj.properties.insert(key, val);
            self.skip_whitespace();
            match self.peek()? {
                b',' => { self.pos += 1; }
                b'}' => { self.pos += 1; break; }
                _ => return None,
            }
        }
        obj.properties.insert(
            "__keys".into(),
            Value::Object(Arc::new(Mutex::new(Object::new_array(tracked_keys)))),
        );
        Some(Value::Object(Arc::new(Mutex::new(obj))))
    }
}

/// UTF-8 leading-byte → sequence length. Returns 0 on invalid.
fn utf8_char_len(b: u8) -> usize {
    if b & 0x80 == 0 { 1 }
    else if b & 0xE0 == 0xC0 { 2 }
    else if b & 0xF0 == 0xE0 { 3 }
    else if b & 0xF8 == 0xF0 { 4 }
    else { 0 }
}
