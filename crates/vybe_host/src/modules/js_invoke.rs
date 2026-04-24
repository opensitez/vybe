//! # `vybe:js-value` — polymorphic method-dispatch shim
//!
//! The wasm js-builtins proposals publish **separate** import modules per
//! reflector (`vybe:js-array.*`, `wasm:js-string.*`, `vybe:js-map.*`,
//! …), which is correct when the compiler knows the receiver's type. For
//! dynamically-typed languages (JS, Python, Ruby) the receiver type isn't
//! known at compile time, so the compiler emits a single dispatch point
//! and defers the method lookup to runtime.
//!
//! This module registers `vybe:js-value.invokeMethod(receiver, name,
//! ...args)`. On v8 via the js-builtins glue, the equivalent shim is one
//! line of JS: `receiver[name](...args)` — native `String.prototype` vs
//! `Array.prototype` dispatch, same prototype-chain walk, same
//! method-missing behaviour. Vybe's in-VM handler mirrors that for the
//! built-in types the runtime knows about and walks the user object's
//! property bag for everything else.
//!
//! # Protocol
//!
//! Stack at the CALL_IMPORT site:
//! * `args[0]`  — receiver (any Value)
//! * `args[1]`  — method name (String)
//! * `args[2..]` — user-supplied arguments
//!
//! Returns the method's result, or `Value::Undefined` if the receiver
//! has no such method (matches JS `TypeError: x.foo is not a function`
//! — we return undefined rather than trap so polyfill-backed code keeps
//! running).

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

fn make_array(elems: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elems))))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "vybe:js-value",
        "invokeMethod",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let method = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Undefined,
            };
            let user_args: &[Value] = if args.len() > 2 { &args[2..] } else { &[] };
            dispatch(ctx, &receiver, &method, user_args)
        }),
    );
}

fn dispatch(ctx: &mut HostContext, receiver: &Value, method: &str, args: &[Value]) -> Value {
    match receiver {
        Value::String(_) => dispatch_string(receiver, method, args),
        Value::Object(obj) => {
            let kind_tag = {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(_) => 1,
                    ObjectKind::Map(_) => 2,
                    ObjectKind::Set(_) => 3,
                    _ => 0,
                }
            };
            match kind_tag {
                1 => dispatch_array(ctx, obj.clone(), method, args),
                2 => dispatch_map(ctx, obj.clone(), method, args),
                3 => dispatch_set(ctx, obj.clone(), method, args),
                _ => dispatch_plain_object(ctx, obj.clone(), method, args),
            }
        }
        _ => Value::Undefined,
    }
}

// ── String methods (`String.prototype.*`) ─────────────────────────────

fn dispatch_string(receiver: &Value, method: &str, args: &[Value]) -> Value {
    let s = match receiver {
        Value::String(s) => s.clone(),
        _ => return Value::Undefined,
    };
    match method {
        "slice" | "substring" => {
            let chars: Vec<char> = s.chars().collect();
            let len = chars.len() as i32;
            let start = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let end = args
                .get(1)
                .and_then(|v| match v {
                    Value::Null | Value::Undefined => None,
                    _ => Some(v.as_i32()),
                })
                .unwrap_or(len);
            let mut s_idx = if method == "substring" {
                start.max(0).min(len) as usize
            } else if start < 0 {
                ((len + start).max(0)) as usize
            } else {
                (start as usize).min(chars.len())
            };
            let mut e_idx = if method == "substring" {
                end.max(0).min(len) as usize
            } else if end < 0 {
                ((len + end).max(0)) as usize
            } else {
                (end as usize).min(chars.len())
            };
            if method == "substring" && s_idx > e_idx {
                std::mem::swap(&mut s_idx, &mut e_idx);
            }
            let out: String = if s_idx < e_idx {
                chars[s_idx..e_idx].iter().collect()
            } else {
                String::new()
            };
            Value::String(Arc::from(out.as_str()))
        }
        "includes" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            let from = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let hay: String = s.chars().skip(from).collect();
            Value::Bool(hay.contains(needle.as_str()))
        }
        "indexOf" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            let from = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let hay: String = s.chars().skip(from).collect();
            match hay.find(needle.as_str()) {
                Some(byte_idx) => {
                    let cp = hay[..byte_idx].chars().count();
                    Value::I32((from + cp) as i32)
                }
                None => Value::I32(-1),
            }
        }
        "lastIndexOf" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            match s.rfind(needle.as_str()) {
                Some(byte_idx) => Value::I32(s[..byte_idx].chars().count() as i32),
                None => Value::I32(-1),
            }
        }
        "startsWith" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            Value::Bool(s.starts_with(needle.as_str()))
        }
        "endsWith" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            Value::Bool(s.ends_with(needle.as_str()))
        }
        "at" => {
            let chars: Vec<char> = s.chars().collect();
            let len = chars.len() as i32;
            let i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let idx = if i < 0 { len + i } else { i };
            if idx < 0 || idx >= len {
                Value::Undefined
            } else {
                Value::String(Arc::from(chars[idx as usize].to_string().as_str()))
            }
        }
        "charAt" => {
            let chars: Vec<char> = s.chars().collect();
            let i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 || (i as usize) >= chars.len() {
                Value::String(Arc::from(""))
            } else {
                Value::String(Arc::from(chars[i as usize].to_string().as_str()))
            }
        }
        "charCodeAt" => {
            let i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            s.chars()
                .nth(i as usize)
                .map(|c| Value::I32(c as i32))
                .unwrap_or(Value::F64(f64::NAN))
        }
        "toUpperCase" => Value::String(Arc::from(s.to_uppercase().as_str())),
        "toLowerCase" => Value::String(Arc::from(s.to_lowercase().as_str())),
        "trim" => Value::String(Arc::from(s.trim())),
        "trimStart" | "trimLeft" => Value::String(Arc::from(s.trim_start())),
        "trimEnd" | "trimRight" => Value::String(Arc::from(s.trim_end())),
        "repeat" => {
            let n = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            Value::String(Arc::from(s.repeat(n).as_str()))
        }
        "split" => {
            let sep = args.first().map(to_str).unwrap_or_default();
            let parts: Vec<Value> = if sep.is_empty() {
                s.chars()
                    .map(|c| Value::String(Arc::from(c.to_string().as_str())))
                    .collect()
            } else {
                s.split(sep.as_str())
                    .map(|p| Value::String(Arc::from(p)))
                    .collect()
            };
            make_array(parts)
        }
        "replace" => {
            let find = args.first().map(to_str).unwrap_or_default();
            let with = args.get(1).map(to_str).unwrap_or_default();
            Value::String(Arc::from(s.replacen(find.as_str(), with.as_str(), 1).as_str()))
        }
        "replaceAll" => {
            let find = args.first().map(to_str).unwrap_or_default();
            let with = args.get(1).map(to_str).unwrap_or_default();
            Value::String(Arc::from(s.replace(find.as_str(), with.as_str()).as_str()))
        }
        "concat" => {
            let mut out = s.to_string();
            for a in args {
                out.push_str(&to_str(a));
            }
            Value::String(Arc::from(out.as_str()))
        }
        "padStart" => pad(&s, args, true),
        "padEnd" => pad(&s, args, false),
        "toString" | "valueOf" => Value::String(s),
        _ => Value::Undefined,
    }
}

fn pad(s: &str, args: &[Value], start: bool) -> Value {
    let target = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
    let pad_char = args
        .get(1)
        .map(to_str)
        .filter(|p| !p.is_empty())
        .unwrap_or_else(|| " ".to_string());
    let cur_len = s.chars().count();
    if cur_len >= target {
        return Value::String(Arc::from(s));
    }
    let needed = target - cur_len;
    let mut pad_str = String::new();
    while pad_str.chars().count() < needed {
        pad_str.push_str(&pad_char);
    }
    let pad_trimmed: String = pad_str.chars().take(needed).collect();
    let out = if start {
        format!("{}{}", pad_trimmed, s)
    } else {
        format!("{}{}", s, pad_trimmed)
    };
    Value::String(Arc::from(out.as_str()))
}

// ── Array methods (`Array.prototype.*`) ──────────────────────────────

fn dispatch_array(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    match method {
        "length" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                Value::I32(v.len() as i32)
            } else {
                Value::I32(0)
            }
        }
        "push" => {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                for a in args {
                    v.push(a.clone());
                }
                sync_length(&mut o);
                return Value::I32(v_len_after(&o));
            }
            Value::I32(0)
        }
        "pop" => {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                let popped = v.pop().unwrap_or(Value::Undefined);
                sync_length(&mut o);
                return popped;
            }
            Value::Undefined
        }
        "shift" => {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                if v.is_empty() {
                    return Value::Undefined;
                }
                let r = v.remove(0);
                sync_length(&mut o);
                return r;
            }
            Value::Undefined
        }
        "unshift" => {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                for (i, a) in args.iter().enumerate() {
                    v.insert(i, a.clone());
                }
                sync_length(&mut o);
                return Value::I32(v_len_after(&o));
            }
            Value::I32(0)
        }
        "slice" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let len = v.len() as i32;
                let start = args.first().map(|a| a.as_i32()).unwrap_or(0);
                let end = args.get(1).map(|a| a.as_i32()).unwrap_or(len);
                let s = (if start < 0 { len + start } else { start }).max(0).min(len) as usize;
                let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                let out: Vec<Value> = if s < e { v[s..e].to_vec() } else { Vec::new() };
                return make_array(out);
            }
            make_array(Vec::new())
        }
        "concat" => {
            let mut out = {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => v.clone(),
                    _ => Vec::new(),
                }
            };
            for a in args {
                match a {
                    Value::Object(other) => {
                        let lo = other.lock().unwrap();
                        match &lo.kind {
                            ObjectKind::Array(v) => out.extend(v.iter().cloned()),
                            _ => out.push(a.clone()),
                        }
                    }
                    _ => out.push(a.clone()),
                }
            }
            make_array(out)
        }
        "includes" => {
            let needle = args.first().cloned().unwrap_or(Value::Undefined);
            let from = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                for elem in v.iter().skip(from) {
                    if elem.eq(&needle) {
                        return Value::Bool(true);
                    }
                }
            }
            Value::Bool(false)
        }
        "indexOf" => {
            let needle = args.first().cloned().unwrap_or(Value::Undefined);
            let from = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                for (i, elem) in v.iter().enumerate().skip(from) {
                    if elem.eq(&needle) {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }
        "lastIndexOf" => {
            let needle = args.first().cloned().unwrap_or(Value::Undefined);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                for (i, elem) in v.iter().enumerate().rev() {
                    if elem.eq(&needle) {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }
        "at" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let len = v.len() as i32;
                let i = args.first().map(|a| a.as_i32()).unwrap_or(0);
                let idx = if i < 0 { len + i } else { i };
                if idx < 0 || idx >= len {
                    return Value::Undefined;
                }
                return v.get(idx as usize).cloned().unwrap_or(Value::Undefined);
            }
            Value::Undefined
        }
        "join" => {
            let sep = args.first().map(to_str).unwrap_or_else(|| ",".to_string());
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let parts: Vec<String> = v
                    .iter()
                    .map(|e| match e {
                        Value::Null | Value::Undefined => String::new(),
                        other => format!("{}", other),
                    })
                    .collect();
                return Value::String(Arc::from(parts.join(&sep).as_str()));
            }
            Value::String(Arc::from(""))
        }
        "reverse" => {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                v.reverse();
            }
            drop(o);
            Value::Object(obj)
        }
        "fill" => {
            let fill = args.first().cloned().unwrap_or(Value::Undefined);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                let len = v.len() as i32;
                let start = args.get(1).map(|a| a.as_i32()).unwrap_or(0);
                let end = args.get(2).map(|a| a.as_i32()).unwrap_or(len);
                let s = (if start < 0 { len + start } else { start }).max(0).min(len) as usize;
                let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                for i in s..e {
                    v[i] = fill.clone();
                }
            }
            drop(o);
            Value::Object(obj)
        }
        "splice" => {
            let start = args.first().map(|a| a.as_i32()).unwrap_or(0);
            let del = args.get(1).map(|a| a.as_i32().max(0) as usize).unwrap_or(0);
            let items: Vec<Value> = args.iter().skip(2).cloned().collect();
            let mut deleted = Vec::new();
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                let len = v.len();
                let idx = if start < 0 {
                    ((len as i32) + start).max(0) as usize
                } else {
                    (start as usize).min(len)
                };
                let end = (idx + del).min(len);
                for _ in idx..end {
                    deleted.push(v.remove(idx));
                }
                for (i, it) in items.into_iter().enumerate() {
                    v.insert(idx + i, it);
                }
                sync_length(&mut o);
            }
            make_array(deleted)
        }
        "forEach" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Undefined,
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            for (i, v) in snapshot.into_iter().enumerate() {
                ctx.invoke(
                    &cb,
                    &[v, Value::I32(i as i32), Value::Object(obj.clone())],
                );
            }
            Value::Undefined
        }
        "map" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return make_array(Vec::new()),
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            let out: Vec<Value> = snapshot
                .into_iter()
                .enumerate()
                .map(|(i, v)| {
                    ctx.invoke(
                        &cb,
                        &[v, Value::I32(i as i32), Value::Object(obj.clone())],
                    )
                })
                .collect();
            make_array(out)
        }
        "filter" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return make_array(Vec::new()),
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            let mut out = Vec::new();
            for (i, v) in snapshot.into_iter().enumerate() {
                let keep = ctx.invoke(
                    &cb,
                    &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())],
                );
                if truthy(&keep) {
                    out.push(v);
                }
            }
            make_array(out)
        }
        "reduce" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Undefined,
            };
            let has_initial = args.len() > 1;
            let mut acc = if has_initial {
                args[1].clone()
            } else {
                Value::Undefined
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            let mut iter = snapshot.into_iter().enumerate();
            if !has_initial {
                if let Some((_, first)) = iter.next() {
                    acc = first;
                }
            }
            for (i, v) in iter {
                acc = ctx.invoke(
                    &cb,
                    &[
                        acc,
                        v,
                        Value::I32(i as i32),
                        Value::Object(obj.clone()),
                    ],
                );
            }
            acc
        }
        "some" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Bool(false),
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            for (i, v) in snapshot.into_iter().enumerate() {
                if truthy(&ctx.invoke(
                    &cb,
                    &[v, Value::I32(i as i32), Value::Object(obj.clone())],
                )) {
                    return Value::Bool(true);
                }
            }
            Value::Bool(false)
        }
        "every" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Bool(true),
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            for (i, v) in snapshot.into_iter().enumerate() {
                if !truthy(&ctx.invoke(
                    &cb,
                    &[v, Value::I32(i as i32), Value::Object(obj.clone())],
                )) {
                    return Value::Bool(false);
                }
            }
            Value::Bool(true)
        }
        "find" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Undefined,
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            for (i, v) in snapshot.into_iter().enumerate() {
                if truthy(&ctx.invoke(
                    &cb,
                    &[
                        v.clone(),
                        Value::I32(i as i32),
                        Value::Object(obj.clone()),
                    ],
                )) {
                    return v;
                }
            }
            Value::Undefined
        }
        "findIndex" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::I32(-1),
            };
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            for (i, v) in snapshot.into_iter().enumerate() {
                if truthy(&ctx.invoke(
                    &cb,
                    &[v, Value::I32(i as i32), Value::Object(obj.clone())],
                )) {
                    return Value::I32(i as i32);
                }
            }
            Value::I32(-1)
        }
        "toString" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let parts: Vec<String> = v
                    .iter()
                    .map(|e| match e {
                        Value::Null | Value::Undefined => String::new(),
                        other => format!("{}", other),
                    })
                    .collect();
                return Value::String(Arc::from(parts.join(",").as_str()));
            }
            Value::String(Arc::from("[object Object]"))
        }
        _ => Value::Undefined,
    }
}

fn v_len_after(o: &Object) -> i32 {
    if let ObjectKind::Array(ref v) = o.kind {
        v.len() as i32
    } else {
        0
    }
}

fn sync_length(o: &mut Object) {
    if let ObjectKind::Array(ref v) = o.kind {
        let n = v.len() as f64;
        o.properties
            .insert("length".to_string(), Value::F64(n));
    }
}

fn truthy(v: &Value) -> bool {
    match v {
        Value::Bool(b) => *b,
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::String(s) => !s.is_empty(),
        Value::Null | Value::Undefined => false,
        _ => true,
    }
}

// ── Map / Set dispatch — delegate to existing host fns where possible.
// Callers on v8 go through `vybe:js-map.*` directly; Vybe's VM does the
// work inline here so it doesn't need `HostContext` to loop back into
// the host registry.

fn dispatch_map(
    _ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    let _ = (obj, method, args);
    // Minimal stub — Map methods are routed through their typed imports at
    // compile time (profile has `[array_methods]` that doesn't intercept
    // Map/Set). If a dynamic-typed call lands here, fall through to
    // Undefined so a useful error surfaces in tests.
    Value::Undefined
}

fn dispatch_set(
    _ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    let _ = (obj, method, args);
    Value::Undefined
}

// ── Plain object / prototype walk ─────────────────────────────────────

fn dispatch_plain_object(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    if method == "hasOwnProperty" {
        let key = args.first().map(to_str).unwrap_or_default();
        let o = obj.lock().unwrap();
        return Value::Bool(o.properties.contains_key(&key));
    }
    // Look up method on the object (callable property).
    let cb = {
        let o = obj.lock().unwrap();
        o.properties.get(method).cloned()
    };
    if let Some(fn_val) = cb {
        if !matches!(fn_val, Value::Null | Value::Undefined) {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            return ctx.invoke(&fn_val, &call_args);
        }
    }
    Value::Undefined
}

// ── Helpers ────────────────────────────────────────────────────────────

fn to_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{}", other),
    }
}
