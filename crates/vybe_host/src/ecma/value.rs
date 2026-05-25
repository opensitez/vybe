//! # `ecma:value` — polymorphic method-dispatch shim
//!
//! The wasm js-builtins proposals publish **separate** import modules per
//! reflector (`ecma:array.*`, `wasm:js-string.*`, `ecma:map.*`,
//! …), which is correct when the compiler knows the receiver's type. For
//! dynamically-typed languages (JS, Python, Ruby) the receiver type isn't
//! known at compile time, so the compiler emits a single dispatch point
//! and defers the method lookup to runtime.
//!
//! This module registers `ecma:value.invokeMethod(receiver, name,
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
use crate::ecma::typedarray::{ta_live_length, read_element, write_element, new_view_over_buffer};
use crate::ecma::weakmap::{WEAKMAP_TAG, WEAKSET_TAG, WM_KEYS_PROP, key_ptr_find as wm_key_ptr_find};

fn make_array(elems: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elems))))
}

pub fn register(vm: &mut VM) {
    // ECMA-262 §13.15.4 Application of `+` operator. For Objects we
    // call their `valueOf` / `toString` per ToPrimitive (§7.1.1) so
    // class instances with a `.toString()` override stringify
    // correctly via `"" + obj`. The VM's DYN_ADD opcode falls back to
    // Display for Objects which yields `[object]` — non-spec for JS.
    vm.register_host_fn(
        "ecma:value",
        "add",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = args.first().cloned().unwrap_or(Value::Undefined);
            let b = args.get(1).cloned().unwrap_or(Value::Undefined);
            let pa = to_primitive(ctx, &a, "default");
            let pb = to_primitive(ctx, &b, "default");
            match (&pa, &pb) {
                (Value::String(_), _) | (_, Value::String(_)) => {
                    Value::String(Arc::from(format!("{}{}", pa, pb).as_str()))
                }
                _ => {
                    let na = pa.as_f64();
                    let nb = pb.as_f64();
                    Value::F64(na + nb)
                }
            }
        }),
    );

    // ECMA-262 §7.1.4 ToNumber. For Object operands runs ToPrimitive
    // with hint "number" then coerces — that's what makes `Date - Date`
    // produce a millisecond delta (Date.prototype.valueOf returns
    // __time) rather than NaN. Plain primitives get the same coercion
    // the VM's `as_f64` already does.
    vm.register_host_fn(
        "ecma:value",
        "toNumber",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            let p = to_primitive(ctx, &v, "number");
            Value::F64(p.as_f64())
        }),
    );

    // ECMA-262 §7.1.1 ToPrimitive(hint=number). Used by the relational
    // operators (`<`, `>`, `<=`, `>=`) before falling through to
    // DYN_LT / DYN_GT / DYN_LE / DYN_GE — those handle string-string
    // lex compare and numeric compare on primitives, but bottom out
    // on `as_f64` (NaN) for Object operands. Pre-coercion makes
    // `Date < Date` and `valueOfObj < n` work without changing the VM.
    vm.register_host_fn(
        "ecma:value",
        "toPrimitive",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            to_primitive(ctx, &v, "number")
        }),
    );

    // Runtime predicate for the JS `.next()` compile path: returns
    // true when the value is an `ObjectKind::Continuation` so the
    // compiler can route through `Op::GEN_NEXT` (WASM stack-switching)
    // for actual generators while keeping the polymorphic dispatch
    // for user-defined `.next()` methods on custom iterables.
    vm.register_host_fn(
        "ecma:value",
        "isGenerator",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let is_cont = matches!(args.first(), Some(Value::Object(o))
                if matches!(o.lock().unwrap().kind,
                    vybe_bytecode::value::ObjectKind::Continuation(_)));
            Value::Bool(is_cont)
        }),
    );

    vm.register_host_fn(
        "ecma:value",
        "isGeneratorDone",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let is_done = match args.first() {
                Some(Value::Object(o)) => {
                    let obj = o.lock().unwrap();
                    match &obj.kind {
                        vybe_bytecode::value::ObjectKind::Continuation(cs) => {
                            matches!(*cs.state.lock().unwrap(), vybe_bytecode::value::ContinuationPhase::Done)
                        }
                        _ => false,
                    }
                }
                _ => false,
            };
            Value::Bool(is_done)
        }),
    );

    // ECMA-262 §13.5.3 Table 41 typeof — returns "object" for both
    // plain objects AND arrays (the VM's REF_TYPEOF opcode reports
    // "array" which is non-spec). Used by the JS compiler so all
    // Vybe outputs match v8/SpiderMonkey/QuickJS behaviour.
    vm.register_host_fn(
        "ecma:value",
        "typeof",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            let tag = match &v {
                Value::Undefined => "undefined",
                Value::Null => "object",
                Value::Bool(_) => "boolean",
                Value::I32(_) | Value::I64(_) | Value::F64(_) => "number",
                Value::String(_) => "string",
                Value::Symbol(_) => "symbol",
                Value::BigInt(_) => "bigint",
                Value::V128(_) => "v128",
                Value::WeakRef(_) => "object",
                Value::Object(o) => {
                    let ob = o.lock().unwrap();
                    match &ob.kind {
                        ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "function",
                        // Spec: arrays are "object", not "array".
                        _ => "object",
                    }
                }
            };
            Value::String(Arc::from(tag))
        }),
    );

    vm.register_host_fn(
        "ecma:value",
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

    vm.register_host_fn(
        "ecma:value",
        "getMethodForCall",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let method = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Undefined,
            };
            lookup_method_for_call(&receiver, &method)
        }),
    );

    vm.register_host_fn(
        "ecma:value",
        "instanceOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let receiver = args.first().cloned().unwrap_or(Value::Undefined);
            let ctor = args.get(1).cloned().unwrap_or(Value::Undefined);
            Value::Bool(js_instanceof(&receiver, &ctor))
        }),
    );
}

fn dispatch(ctx: &mut HostContext, receiver: &Value, method: &str, args: &[Value]) -> Value {
    match receiver {
        Value::F64(_) | Value::I32(_) | Value::I64(_) => dispatch_number(receiver, method, args),
        Value::String(_) => dispatch_string(ctx, receiver, method, args),
        Value::Symbol(desc) => match method {
            // ECMA-262 §20.4.3.3 Symbol.prototype.toString — "Symbol(<desc>)"
            "toString" => Value::String(Arc::from(format!("Symbol({})", desc).as_str())),
            // ECMA-262 §20.4.3.4 Symbol.prototype.valueOf — returns the symbol itself
            "valueOf" => receiver.clone(),
            // ECMA-262 §20.4.3.2 Symbol.prototype.description — the raw description string
            "description" => {
                if desc.is_empty() {
                    Value::Undefined
                } else {
                    Value::String(Arc::clone(desc))
                }
            }
            _ => Value::Undefined,
        },
        Value::BigInt(n) => dispatch_bigint(*n, method, args),
        Value::Object(obj) => {
            // WeakMap/WeakSet use ObjectKind::Array backing — check their tag before kind dispatch.
            let kind_tag = {
                let o = obj.lock().unwrap();
                if o.properties.contains_key(WEAKMAP_TAG) { 5 }
                else if o.properties.contains_key(WEAKSET_TAG) { 6 }
                else {
                    match &o.kind {
                        ObjectKind::Array(_)       => 1,
                        ObjectKind::Map(_)         => 2,
                        ObjectKind::Set(_)         => 3,
                        ObjectKind::TypedArray(_)  => 4,
                        ObjectKind::ArrayBuffer(_) => 7,
                        _ => {
                            if o.properties.contains_key(crate::ecma::arraybuffer::DV_TAG) { 8 }
                            else { 0 }
                        }
                    }
                }
            };
            match kind_tag {
                1 => dispatch_array(ctx, obj.clone(), method, args),
                2 => dispatch_map(ctx, obj.clone(), method, args),
                3 => dispatch_set(ctx, obj.clone(), method, args),
                4 => dispatch_typed_array(ctx, obj.clone(), method, args),
                5 => dispatch_weakmap(obj.clone(), method, args),
                6 => dispatch_weakset(obj.clone(), method, args),
                7 => crate::ecma::arraybuffer::dispatch_arraybuffer_method(obj.clone(), method, args)
                        .unwrap_or_else(|| dispatch_plain_object(ctx, obj.clone(), method, args)),
                8 => crate::ecma::arraybuffer::dispatch_dataview_method(obj.clone(), method, args)
                        .unwrap_or_else(|| dispatch_plain_object(ctx, obj.clone(), method, args)),
                _ => dispatch_plain_object(ctx, obj.clone(), method, args),
            }
        }
        _ => Value::Undefined,
    }
}

fn dispatch_bigint(n: i64, method: &str, args: &[Value]) -> Value {
    match method {
        "toString" => {
            let radix = args.first().map(|v| v.as_i32() as u32).unwrap_or(10);
            if radix == 10 || radix < 2 || radix > 36 {
                return Value::String(Arc::from(format!("{}", n).as_str()));
            }
            let negative = n < 0;
            let mut v = (n as i128).unsigned_abs();
            if v == 0 { return Value::String(Arc::from("0")); }
            let mut out = String::new();
            while v > 0 {
                let digit = (v % radix as u128) as u32;
                out.insert(0, char::from_digit(digit, radix).unwrap_or('?'));
                v /= radix as u128;
            }
            if negative { out.insert(0, '-'); }
            Value::String(Arc::from(out.as_str()))
        }
        "valueOf" => Value::BigInt(n),
        _ => Value::Undefined,
    }
}

// ── Number methods (`Number.prototype.*`) ─────────────────────────────

fn dispatch_number(receiver: &Value, method: &str, args: &[Value]) -> Value {
    let n = receiver.as_f64();
    match method {
        "toString" => {
            let radix = args.first().map(|v| v.as_i32() as u32).unwrap_or(10);
            if radix == 10 || radix < 2 || radix > 36 {
                if n.is_finite() && n.fract() == 0.0 {
                    return Value::String(Arc::from(format!("{}", n as i64).as_str()));
                }
                return Value::String(Arc::from(format!("{}", n).as_str()));
            }
            let int_val = n as i64;
            let negative = int_val < 0;
            let mut v = (int_val as i128).unsigned_abs();
            if v == 0 { return Value::String(Arc::from("0")); }
            let mut out = String::new();
            while v > 0 {
                let digit = (v % radix as u128) as u32;
                out.insert(0, char::from_digit(digit, radix).unwrap_or('?'));
                v /= radix as u128;
            }
            if negative { out.insert(0, '-'); }
            Value::String(Arc::from(out.as_str()))
        }
        "toFixed" => {
            let digits = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            Value::String(Arc::from(format!("{:.1$}", n, digits).as_str()))
        }
        "toExponential" => {
            let digits = args.first().map(|v| v.as_i32().max(0) as usize).unwrap_or(6);
            let raw = format!("{:.1$e}", n, digits);
            // Rust uses e4 but JS uses e+4; normalize
            let parts: Vec<&str> = raw.splitn(2, 'e').collect();
            if parts.len() == 2 {
                let exp: i32 = parts[1].parse().unwrap_or(0);
                let sign = if exp >= 0 { "+" } else { "" };
                Value::String(Arc::from(format!("{}e{}{}", parts[0], sign, exp).as_str()))
            } else {
                Value::String(Arc::from(raw.as_str()))
            }
        }
        "toPrecision" => {
            let prec = args.first().map(|v| v.as_i32().max(1) as usize).unwrap_or(0);
            if prec == 0 { return Value::String(Arc::from(format!("{}", n).as_str())); }
            Value::String(Arc::from(format!("{:.prec$}", n, prec = prec).as_str()))
        }
        "toLocaleString" => {
            if n.is_finite() && n.fract() == 0.0 {
                Value::String(Arc::from(format!("{}", n as i64).as_str()))
            } else {
                Value::String(Arc::from(format!("{}", n).as_str()))
            }
        }
        "valueOf" => receiver.clone(),
        _ => Value::Undefined,
    }
}

// ── String methods (`String.prototype.*`) ─────────────────────────────

fn dispatch_string(ctx: &mut HostContext, receiver: &Value, method: &str, args: &[Value]) -> Value {
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
            let chars: Vec<char> = s.chars().collect();
            let len = chars.len() as i32;
            let from_idx = args.get(1)
                .and_then(|v| match v { Value::Undefined | Value::Null => None, _ => Some(v.as_i32()) })
                .unwrap_or(len);
            let from_idx = (from_idx.max(0) as usize).min(chars.len());
            let search_end = (from_idx + needle.chars().count()).min(chars.len());
            let hay: String = chars[..search_end].iter().collect();
            match hay.rfind(needle.as_str()) {
                Some(byte_idx) => Value::I32(hay[..byte_idx].chars().count() as i32),
                None => Value::I32(-1),
            }
        }
        "startsWith" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            let chars: Vec<char> = s.chars().collect();
            let pos = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0).min(chars.len());
            let hay: String = chars[pos..].iter().collect();
            Value::Bool(hay.starts_with(needle.as_str()))
        }
        "endsWith" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            let chars: Vec<char> = s.chars().collect();
            let end_pos = args.get(1)
                .and_then(|v| match v { Value::Undefined | Value::Null => None, _ => Some(v.as_i32()) })
                .map(|n| (n.max(0) as usize).min(chars.len()))
                .unwrap_or(chars.len());
            let hay: String = chars[..end_pos].iter().collect();
            Value::Bool(hay.ends_with(needle.as_str()))
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
            // ECMA-262 §22.1.3.20 — first arg can be a String OR a RegExp.
            // Detect the RegExp shape (object stamped __type=RegExp) and
            // dispatch through `ecma:regexp` for shared regex semantics.
            if let Some((pat, flags)) = regex_pattern(args.first()) {
                let mut call_args = Vec::with_capacity(args.len() + 1);
                call_args.push(Value::String(s.clone()));
                let _ = (pat, flags);
                call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
                call_args.extend_from_slice(&args[1..]);
                if let Some(result) = crate::ecma::regexp::dispatch_regexp_string_method(ctx, "split", &call_args) {
                    return result;
                }
            }
            let sep = args.first().map(to_str).unwrap_or_default();
            let limit = args.get(1).and_then(|v| {
                let n = v.as_i32();
                if n > 0 { Some(n as usize) } else { None }
            });
            let parts: Vec<Value> = if sep.is_empty() {
                let chars = s.chars()
                    .map(|c| Value::String(Arc::from(c.to_string().as_str())));
                match limit {
                    Some(n) => chars.take(n).collect(),
                    None => chars.collect(),
                }
            } else {
                let pieces = s.split(sep.as_str())
                    .map(|p| Value::String(Arc::from(p)));
                match limit {
                    Some(n) => pieces.take(n).collect(),
                    None => pieces.collect(),
                }
            };
            make_array(parts)
        }
        "replace" => {
            // ECMA-262 §22.1.3.18 — first arg can be String or RegExp.
            // With a RegExp + `g` flag → replace all; else first only.
            // Replacement may be a callable (function or host fn); per
            // spec the function is called with (match, ...captures, offset, input)
            // and its return value is the substitution.
            let replacement = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some((pat, flags)) = regex_pattern(args.first()) {
                let mut call_args = Vec::with_capacity(3);
                call_args.push(Value::String(s.clone()));
                let _ = (pat, flags);
                call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
                call_args.push(replacement.clone());
                if let Some(result) = crate::ecma::regexp::dispatch_regexp_string_method(ctx, "replace", &call_args) {
                    return result;
                }
            }
            let is_callable = matches!(&replacement, Value::Object(o)
                if matches!(o.lock().unwrap().kind,
                    vybe_bytecode::value::ObjectKind::Function(_)
                    | vybe_bytecode::value::ObjectKind::HostFunction(_)));
            let find = args.first().map(to_str).unwrap_or_default();
            if is_callable {
                // Plain-string find with callable replacement: replace
                // first occurrence by invoking the callback once.
                let result = match s.find(find.as_str()) {
                    Some(pos) => {
                        let cb_args = vec![
                            Value::String(Arc::from(find.as_str())),
                            Value::I32(pos as i32),
                            Value::String(s.clone()),
                        ];
                        let ret = ctx.invoke(&replacement, &cb_args);
                        let with = match ret {
                            Value::String(ref st) => st.to_string(),
                            other => format!("{}", other),
                        };
                        format!("{}{}{}", &s[..pos], with, &s[pos + find.len()..])
                    }
                    None => s.to_string(),
                };
                return Value::String(Arc::from(result.as_str()));
            }
            let with = to_str(&replacement);
            Value::String(Arc::from(s.replacen(find.as_str(), with.as_str(), 1).as_str()))
        }
        "replaceAll" => {
            if let Some((pat, flags)) = regex_pattern(args.first()) {
                let mut call_args = Vec::with_capacity(3);
                call_args.push(Value::String(s.clone()));
                let _ = (pat, flags);
                call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
                call_args.push(args.get(1).cloned().unwrap_or(Value::Undefined));
                if let Some(result) = crate::ecma::regexp::dispatch_regexp_string_method(ctx, "replaceAll", &call_args) {
                    return result;
                }
            }
            let find = args.first().map(to_str).unwrap_or_default();
            let replacement = args.get(1).cloned().unwrap_or(Value::Undefined);
            let is_callable = matches!(&replacement, Value::Object(o)
                if matches!(o.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)));
            if is_callable && !find.is_empty() {
                let mut result = String::new();
                let mut rest = s.as_ref();
                let mut offset = 0usize;
                while let Some(pos) = rest.find(find.as_str()) {
                    result.push_str(&rest[..pos]);
                    let matched = &rest[pos..pos + find.len()];
                    let cb_result = ctx.invoke(&replacement, &[
                        Value::String(Arc::from(matched)),
                        Value::I32((offset + pos) as i32),
                        Value::String(s.clone()),
                    ]);
                    result.push_str(&to_str(&cb_result));
                    offset += pos + find.len();
                    rest = &rest[pos + find.len()..];
                }
                result.push_str(rest);
                return Value::String(Arc::from(result.as_str()));
            }
            let with = to_str(&replacement);
            Value::String(Arc::from(s.replace(find.as_str(), with.as_str()).as_str()))
        }
        "match" => {
            // ECMA-262 §22.1.3.13 — receiver=string, arg=regex (or string,
            // which is treated as a regex source).
            let mut call_args = Vec::with_capacity(2);
            call_args.push(Value::String(s.clone()));
            call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
            if let Some(result) = crate::ecma::regexp::dispatch_regexp_string_method(ctx, "match", &call_args) {
                result
            } else {
                Value::Null
            }
        }
        "search" => {
            let mut call_args = Vec::with_capacity(2);
            call_args.push(Value::String(s.clone()));
            call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
            match crate::ecma::regexp::dispatch_regexp_string_method(ctx, "search", &call_args) {
                Some(result) => result,
                None => Value::I32(-1),
            }
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
        "copyWithin" => {
            let target = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            let end = args.get(2).map(|v| v.as_i32()).unwrap_or(i32::MAX);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                let len = v.len() as i32;
                let t = target.max(0).min(len) as usize;
                let s = start.max(0).min(len) as usize;
                let e = end.max(0).min(len) as usize;
                let slice: Vec<Value> = v[s..e].iter().cloned().collect();
                let max_copy = (len as usize - t).min(slice.len());
                v[t..t + max_copy].clone_from_slice(&slice[..max_copy]);
                sync_length(&mut o);
            }
            drop(o);
            Value::Object(obj)
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
                let len = v.len() as i32;
                let from = args.get(1)
                    .and_then(|x| match x { Value::Undefined | Value::Null => None, _ => Some(x.as_i32()) })
                    .map(|n| if n < 0 { (len + n).max(0) as usize } else { n.min(len - 1).max(0) as usize })
                    .unwrap_or(v.len().saturating_sub(1));
                for i in (0..=from.min(v.len().saturating_sub(1))).rev() {
                    if v[i].eq(&needle) {
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
        "keys" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let out: Vec<Value> = (0..v.len()).map(|i| Value::F64(i as f64)).collect();
                return crate::ecma::array::make_array_iterator(out);
            }
            crate::ecma::array::make_array_iterator(Vec::new())
        }
        "values" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                return crate::ecma::array::make_array_iterator(v.clone());
            }
            crate::ecma::array::make_array_iterator(Vec::new())
        }
        "entries" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let out: Vec<Value> = v
                    .iter()
                    .enumerate()
                    .map(|(i, e)| make_array(vec![Value::F64(i as f64), e.clone()]))
                    .collect();
                return crate::ecma::array::make_array_iterator(out);
            }
            crate::ecma::array::make_array_iterator(Vec::new())
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

// ── Map / Set dispatch — operate directly on ObjectKind::Map/Set.
//
// Callers on v8 go through `ecma:map.*` / `ecma:set.*` directly; Vybe's VM
// does the same work inline here so it doesn't need to loop back through
// `HostContext` into the host registry. Semantics mirror
// `crates/vybe_host/src/ecma/{map,set}.rs` exactly — keys use SameValueZero
// (`Value`'s `Hash + Eq` impl), `delete` uses `shift_remove` to preserve
// insertion order per ECMA-262 §24.1.3.3 / §24.2.3.4.

fn sync_map_size(o: &mut Object) {
    if let ObjectKind::Map(ref m) = o.kind {
        let n = m.len() as i32;
        o.properties.insert("size".to_string(), Value::I32(n));
    }
}

fn sync_set_size(o: &mut Object) {
    if let ObjectKind::Set(ref s) = o.kind {
        let n = s.len() as i32;
        o.properties.insert("size".to_string(), Value::I32(n));
    }
}

fn dispatch_map(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    match method {
        "get" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                return im.get(&key).cloned().unwrap_or(Value::Undefined);
            }
            Value::Undefined
        }
        "set" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            let val = args.get(1).cloned().unwrap_or(Value::Undefined);
            {
                let mut m = obj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = m.kind {
                    im.insert(key, val);
                }
                sync_map_size(&mut m);
            }
            Value::Object(obj)
        }
        "has" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                return Value::Bool(im.contains_key(&key));
            }
            Value::Bool(false)
        }
        "delete" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            let mut m = obj.lock().unwrap();
            let removed = if let ObjectKind::Map(ref mut im) = m.kind {
                im.shift_remove(&key).is_some()
            } else {
                false
            };
            sync_map_size(&mut m);
            Value::Bool(removed)
        }
        "clear" => {
            let mut m = obj.lock().unwrap();
            if let ObjectKind::Map(ref mut im) = m.kind {
                im.clear();
            }
            sync_map_size(&mut m);
            Value::Undefined
        }
        "size" => {
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                return Value::I32(im.len() as i32);
            }
            Value::I32(0)
        }
        "keys" => {
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                return make_array(im.keys().cloned().collect());
            }
            make_array(Vec::new())
        }
        "values" => {
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                return make_array(im.values().cloned().collect());
            }
            make_array(Vec::new())
        }
        "entries" => {
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                let pairs: Vec<Value> = im
                    .iter()
                    .map(|(k, v)| make_array(vec![k.clone(), v.clone()]))
                    .collect();
                return make_array(pairs);
            }
            make_array(Vec::new())
        }
        "forEach" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<(Value, Value)> = {
                let m = obj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    im.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
                } else {
                    Vec::new()
                }
            };
            for (k, v) in snapshot {
                ctx.invoke(&cb, &[v, k, Value::Object(obj.clone())]);
            }
            Value::Undefined
        }
        _ => Value::Undefined,
    }
}

fn dispatch_set(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    match method {
        "add" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            {
                let mut so = obj.lock().unwrap();
                if let ObjectKind::Set(ref mut s) = so.kind {
                    s.insert(v);
                }
                sync_set_size(&mut so);
            }
            Value::Object(obj)
        }
        "has" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            let so = obj.lock().unwrap();
            if let ObjectKind::Set(ref s) = so.kind {
                return Value::Bool(s.contains(&v));
            }
            Value::Bool(false)
        }
        "delete" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            let mut so = obj.lock().unwrap();
            let removed = if let ObjectKind::Set(ref mut s) = so.kind {
                s.shift_remove(&v)
            } else {
                false
            };
            sync_set_size(&mut so);
            Value::Bool(removed)
        }
        "clear" => {
            let mut so = obj.lock().unwrap();
            if let ObjectKind::Set(ref mut s) = so.kind {
                s.clear();
            }
            sync_set_size(&mut so);
            Value::Undefined
        }
        "size" => {
            let so = obj.lock().unwrap();
            if let ObjectKind::Set(ref s) = so.kind {
                return Value::I32(s.len() as i32);
            }
            Value::I32(0)
        }
        // Set.prototype.keys/values/entries: spec returns an iterator;
        // MVP returns a snapshot Array (matches `ecma:set` registrations).
        "keys" | "values" => {
            let so = obj.lock().unwrap();
            if let ObjectKind::Set(ref s) = so.kind {
                return make_array(s.iter().cloned().collect());
            }
            make_array(Vec::new())
        }
        "entries" => {
            let so = obj.lock().unwrap();
            if let ObjectKind::Set(ref s) = so.kind {
                let pairs: Vec<Value> = s
                    .iter()
                    .map(|v| make_array(vec![v.clone(), v.clone()]))
                    .collect();
                return make_array(pairs);
            }
            make_array(Vec::new())
        }
        "forEach" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<Value> = {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    s.iter().cloned().collect()
                } else {
                    Vec::new()
                }
            };
            for v in snapshot {
                ctx.invoke(&cb, &[v.clone(), v, Value::Object(obj.clone())]);
            }
            Value::Undefined
        }
        _ => Value::Undefined,
    }
}

// ── TypedArray methods (`TypedArray.prototype.*`) ─────────────────────

fn dispatch_typed_array(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    match method {
        "length" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                return Value::I32(ta_live_length(ta) as i32);
            }
            Value::I32(0)
        }
        "indexOf" => {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in 0..live {
                    if read_element(ta, i) == target {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }
        "includes" => {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in 0..live {
                    if read_element(ta, i) == target {
                        return Value::Bool(true);
                    }
                }
            }
            Value::Bool(false)
        }
        "join" => {
            let sep = args.first().map(|v| format!("{v}")).unwrap_or_else(|| ",".to_string());
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let parts: Vec<String> = (0..live).map(|i| format!("{}", read_element(ta, i))).collect();
                return Value::String(Arc::from(parts.join(&sep).as_str()));
            }
            Value::String(Arc::from(""))
        }
        "fill" => {
            let val = args.first().cloned().unwrap_or(Value::Undefined);
            {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    let start = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0).min(live);
                    let end = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(live).min(live);
                    for i in start..end { write_element(ta, i, &val); }
                }
            }
            Value::Object(obj)
        }
        "reverse" => {
            {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    let mut i = 0usize;
                    let mut j = live.saturating_sub(1);
                    while i < j {
                        let a = read_element(ta, i);
                        let b = read_element(ta, j);
                        write_element(ta, i, &b);
                        write_element(ta, j, &a);
                        i += 1; j -= 1;
                    }
                }
            }
            Value::Object(obj)
        }
        "sort" => {
            {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    let mut values: Vec<Value> = (0..live).map(|i| read_element(ta, i)).collect();
                    values.sort_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal));
                    for (i, v) in values.iter().enumerate() { write_element(ta, i, v); }
                }
            }
            Value::Object(obj)
        }
        "slice" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta) as i32;
                let start = args.first().map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(1).map(|v| v.as_i32()).unwrap_or(live);
                let s = start.max(0).min(live) as usize;
                let e = end.max(0).min(live) as usize;
                let elems: Vec<Value> = (s..e).map(|i| read_element(ta, i)).collect();
                return make_array(elems);
            }
            make_array(Vec::new())
        }
        "forEach" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<Value> = {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    (0..live).map(|i| read_element(ta, i)).collect()
                } else { Vec::new() }
            };
            for (i, v) in snapshot.iter().enumerate() {
                ctx.invoke(&cb, &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())]);
            }
            Value::Undefined
        }
        "map" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<Value> = {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    (0..live).map(|i| read_element(ta, i)).collect()
                } else { Vec::new() }
            };
            let out: Vec<Value> = snapshot.iter().enumerate()
                .map(|(i, v)| ctx.invoke(&cb, &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())]))
                .collect();
            make_array(out)
        }
        "filter" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<Value> = {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    (0..live).map(|i| read_element(ta, i)).collect()
                } else { Vec::new() }
            };
            let out: Vec<Value> = snapshot.iter().enumerate()
                .filter(|(i, v)| {
                    let r = ctx.invoke(&cb, &[(*v).clone(), Value::I32(*i as i32), Value::Object(obj.clone())]);
                    matches!(r, Value::Bool(true)) || matches!(r, Value::I32(n) if n != 0)
                })
                .map(|(_, v)| v.clone())
                .collect();
            make_array(out)
        }
        "some" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in 0..live {
                    let v = read_element(ta, i);
                    let r = ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]);
                    if matches!(r, Value::Bool(true)) || matches!(r, Value::I32(n) if n != 0) {
                        return Value::Bool(true);
                    }
                }
            }
            Value::Bool(false)
        }
        "every" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in 0..live {
                    let v = read_element(ta, i);
                    let r = ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]);
                    if !matches!(r, Value::Bool(true)) && !matches!(r, Value::I32(n) if n != 0) {
                        return Value::Bool(false);
                    }
                }
                return Value::Bool(true);
            }
            Value::Bool(true)
        }
        "find" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in 0..live {
                    let v = read_element(ta, i);
                    let r = ctx.invoke(&cb, &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())]);
                    if matches!(r, Value::Bool(true)) || matches!(r, Value::I32(n) if n != 0) {
                        return v;
                    }
                }
            }
            Value::Undefined
        }
        "findIndex" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in 0..live {
                    let v = read_element(ta, i);
                    let r = ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]);
                    if matches!(r, Value::Bool(true)) || matches!(r, Value::I32(n) if n != 0) {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }
        "reduce" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let (mut acc, start) = if args.len() >= 2 {
                    (args[1].clone(), 0)
                } else if live > 0 {
                    (read_element(ta, 0), 1)
                } else {
                    return Value::Undefined;
                };
                for i in start..live {
                    let v = read_element(ta, i);
                    acc = ctx.invoke(&cb, &[acc, v, Value::I32(i as i32), Value::Object(obj.clone())]);
                }
                return acc;
            }
            Value::Undefined
        }
        "lastIndexOf" => {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for i in (0..live).rev() {
                    if Value::same_value_zero(&read_element(ta, i), &target) {
                        return Value::I32(i as i32);
                    }
                }
            }
            Value::I32(-1)
        }
        "set" => {
            // ta.set(source, offset?) — copy elements from source array/TypedArray
            let offset = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let source_values: Vec<Value> = match args.first() {
                Some(Value::Object(src)) => {
                    let s = src.lock().unwrap();
                    match &s.kind {
                        ObjectKind::Array(elems) => elems.clone(),
                        ObjectKind::TypedArray(src_ta) => (0..ta_live_length(src_ta))
                            .map(|i| read_element(src_ta, i))
                            .collect(),
                        _ => Vec::new(),
                    }
                }
                _ => Vec::new(),
            };
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                for (i, v) in source_values.iter().enumerate() {
                    let idx = offset + i;
                    if idx >= live { break; }
                    write_element(ta, idx, v);
                }
            }
            Value::Undefined
        }
        "subarray" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta) as i32;
                let start = args.first().map(|v| v.as_i32()).unwrap_or(0);
                let end = args.get(1).map(|v| v.as_i32()).unwrap_or(live);
                let s = (if start < 0 { live + start } else { start }).max(0).min(live) as usize;
                let e = (if end < 0 { live + end } else { end }).max(0).min(live) as usize;
                let sub_len = if s < e { e - s } else { 0 };
                let buffer_obj = ta.buffer_obj.clone();
                let elem = ta.elem;
                let bpe = elem.bytes_per_element();
                let abs_offset = ta.byte_offset + s * bpe;
                drop(o);
                return new_view_over_buffer(elem, buffer_obj, abs_offset, sub_len);
            }
            Value::Undefined
        }
        "copyWithin" => {
            {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta) as i32;
                    let target = args.first().map(|v| v.as_i32()).unwrap_or(0);
                    let start = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
                    let end = args.get(2).map(|v| v.as_i32()).unwrap_or(live);
                    let t = (if target < 0 { live + target } else { target }).max(0).min(live) as usize;
                    let s = (if start < 0 { live + start } else { start }).max(0).min(live) as usize;
                    let e = (if end < 0 { live + end } else { end }).max(0).min(live) as usize;
                    let snapshot: Vec<Value> = (s..e).map(|i| read_element(ta, i)).collect();
                    let max_copy = (live as usize - t).min(snapshot.len());
                    for (i, v) in snapshot[..max_copy].iter().enumerate() {
                        write_element(ta, t + i, v);
                    }
                }
            }
            Value::Object(obj)
        }
        "keys" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let ks: Vec<Value> = (0..live as i32).map(Value::I32).collect();
                return make_array(ks);
            }
            make_array(Vec::new())
        }
        "values" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let vs: Vec<Value> = (0..live).map(|i| read_element(ta, i)).collect();
                return make_array(vs);
            }
            make_array(Vec::new())
        }
        _ => Value::Undefined,
    }
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
    if method == "propertyIsEnumerable" {
        let key = args.first().map(to_str).unwrap_or_default();
        let o = obj.lock().unwrap();
        let has_own = o.properties.contains_key(&key) && !key.starts_with("__");
        if !has_own { return Value::Bool(false); }
        let is_enum = match o.properties.get("__nonenum") {
            Some(Value::Object(arr)) => {
                let a = arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = a.kind {
                    !elems.iter().any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
                } else { true }
            }
            _ => true,
        };
        return Value::Bool(is_enum);
    }
    // Walk own properties then __proto__ chain for a callable method.
    let cb = {
        let mut found: Option<Value> = None;
        let mut current: Option<Arc<Mutex<Object>>> = Some(obj.clone());
        while let Some(cur) = current {
            let (prop, proto) = {
                let o = cur.lock().unwrap();
                (o.properties.get(method).cloned(), o.properties.get("__proto__").cloned())
            };
            if let Some(v) = prop {
                if !matches!(v, Value::Null | Value::Undefined) {
                    found = Some(v);
                    break;
                }
            }
            current = match proto {
                Some(Value::Object(p)) => Some(p),
                _ => None,
            };
        }
        found
    };
    if let Some(fn_val) = cb {
        return ctx.invoke(&fn_val, args);
    }
    // Type-tagged object fallback: known stamped-`__type` instances
    // (Date) get their methods inline. The polymorphic invokeMethod
    // shim doesn't see the type registry, so `d.toString()` would
    // otherwise return undefined when the instance has no callable
    // `toString` property of its own. ECMA-262 §21.4.4 dispatches
    // these via the Date prototype — same semantics, inline impl.
    let type_tag = {
        let o = obj.lock().unwrap();
        o.properties.get("__type").map(|v| format!("{}", v))
    };
    if let Some(tag) = type_tag {
        if tag == "Date" {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            if let Some(result) = crate::ecma::date::dispatch_date_method(method, &call_args) {
                return result;
            }
        } else if tag == "RegExp" {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            if let Some(result) = crate::ecma::regexp::dispatch_regexp_method(method, &call_args) {
                return result;
            }
        } else if tag == "Promise" {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            if let Some(result) = crate::ecma::promise::dispatch_promise_method(ctx, method, &call_args) {
                return result;
            }
        }
    }
    Value::Undefined
}

fn lookup_method_for_call(receiver: &Value, method: &str) -> Value {
    let Value::Object(receiver_obj) = receiver else {
        return Value::Null;
    };

    let mut current = Some(receiver_obj.clone());
    while let Some(obj) = current {
        let (found, next_proto) = {
            let o = obj.lock().unwrap();
            (
                o.properties.get(method).cloned(),
                o.properties.get("__proto__").cloned(),
            )
        };

        if let Some(value) = found {
            if !matches!(value, Value::Null | Value::Undefined) {
                return bind_method_receiver(receiver_obj.clone(), value);
            }
        }

        current = match next_proto {
            Some(Value::Object(proto)) => Some(proto),
            _ => None,
        };
    }

    Value::Null
}

fn bind_method_receiver(receiver: Arc<Mutex<Object>>, method: Value) -> Value {
    let Value::Object(target) = method else {
        return method;
    };

    let (kind, existing_bound) = {
        let o = target.lock().unwrap();
        match &o.kind {
            ObjectKind::HostFunction(_) => {
                let prev_bound = match o.properties.get("__bound_args") {
                    Some(Value::Object(bound)) => {
                        let bo = bound.lock().unwrap();
                        if let ObjectKind::Array(ref values) = bo.kind {
                            values.clone()
                        } else {
                            Vec::new()
                        }
                    }
                    _ => Vec::new(),
                };
                (Some(o.kind.clone()), prev_bound)
            }
            _ => (None, Vec::new()),
        }
    };

    let Some(kind) = kind else {
        return Value::Object(target);
    };

    let mut combined = Vec::with_capacity(existing_bound.len() + 1);
    combined.push(Value::Object(receiver));
    combined.extend(existing_bound);

    let mut bound_obj = Object::new();
    bound_obj.kind = kind;
    bound_obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(combined)))),
    );
    Value::Object(Arc::new(Mutex::new(bound_obj)))
}

fn js_instanceof(receiver: &Value, ctor: &Value) -> bool {
    let Value::Object(obj) = receiver else {
        return false;
    };

    let ctor_name = match ctor {
        Value::String(name) => Some(name.to_string()),
        Value::Object(ctor_obj) => {
            let ctor_lock = ctor_obj.lock().unwrap();
            match ctor_lock.properties.get("name") {
                Some(Value::String(name)) => Some(name.to_string()),
                Some(other) => Some(format!("{}", other)),
                None => None,
            }
        }
        _ => None,
    };

    if let Some(name) = ctor_name.as_deref() {
        let matched_stamp = {
            let o = obj.lock().unwrap();
            if matches!(o.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == name) {
                true
            } else {
                match o.properties.get("__types") {
                    Some(Value::Object(arr)) => {
                        let arr_lock = arr.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = arr_lock.kind {
                            elems.iter().any(|value| matches!(value, Value::String(tag) if tag.as_ref() == name))
                        } else {
                            false
                        }
                    }
                    _ => false,
                }
            }
        };
        if matched_stamp {
            return true;
        }
    }

    let Value::Object(ctor_obj) = ctor else {
        return false;
    };

    let target_proto = {
        let ctor_lock = ctor_obj.lock().unwrap();
        match ctor_lock.properties.get("prototype") {
            Some(Value::Object(proto)) => Some(proto.clone()),
            _ => None,
        }
    };

    let Some(target_proto) = target_proto else {
        return false;
    };

    let mut current = Some(obj.clone());
    while let Some(cur) = current {
        let next = {
            let lock = cur.lock().unwrap();
            lock.properties.get("__proto__").cloned()
        };
        match next {
            Some(Value::Object(proto)) => {
                if Arc::ptr_eq(&proto, &target_proto) {
                    return true;
                }
                current = Some(proto);
            }
            _ => return false,
        }
    }

    false
}

// ── Helpers ────────────────────────────────────────────────────────────

fn to_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{}", other),
    }
}

/// Walk own + `__proto__` chain looking for a method named `key`.
/// Returns the first non-null value found (so prototype-installed
/// methods are reachable from instances). Bound to 100 hops to
/// protect against accidental cycles.
fn lookup_method_via_proto(obj: &Arc<Mutex<Object>>, key: &str) -> Option<Value> {
    let mut current = obj.clone();
    for _ in 0..100 {
        let next_proto = {
            let o = current.lock().unwrap();
            if let Some(v) = o.properties.get(key) {
                if !matches!(v, Value::Null | Value::Undefined) {
                    return Some(v.clone());
                }
            }
            match o.properties.get("__proto__").cloned() {
                Some(Value::Object(p)) => Some(p),
                _ => None,
            }
        };
        match next_proto {
            Some(p) => current = p,
            None => break,
        }
    }
    None
}

/// ECMA-262 §7.1.1 ToPrimitive — for objects, invoke `valueOf` /
/// `toString` (in `default` / `number` hint order: valueOf then
/// toString; in `string` hint order: toString then valueOf). Returns
/// the first non-Object result; falls back to Display if neither
/// callable yields a primitive.
fn to_primitive(ctx: &mut HostContext, v: &Value, hint: &str) -> Value {
    let obj = match v {
        Value::Object(o) => o.clone(),
        _ => return v.clone(),
    };
    // ECMA-262 §7.1.1: check [Symbol.toPrimitive] first (stored as "toprimitive")
    let tp = obj.lock().unwrap().properties.get("toprimitive").cloned();
    if let Some(tp_fn) = tp {
        if !matches!(tp_fn, Value::Null | Value::Undefined) {
            let hint_val = Value::String(Arc::from(hint));
            return ctx.invoke(&tp_fn, &[v.clone(), hint_val]);
        }
    }
    // Skip the dance for non-Ordinary objects (Functions, Continuations
    // etc. don't have callable valueOf/toString in our model).
    let is_ordinary = matches!(
        obj.lock().unwrap().kind,
        ObjectKind::Ordinary | ObjectKind::Array(_)
    );
    if !is_ordinary {
        return v.clone();
    }
    // Built-in tagged objects: their valueOf / toString live in the
    // dispatch tables (`dispatch_date_method` etc.) rather than the
    // prototype chain. Route through the same channel so `Date - Date`
    // hits ECMA §21.4.4.41 valueOf and yields the ms delta.
    let type_tag = {
        let o = obj.lock().unwrap();
        o.properties.get("__type").and_then(|v| match v {
            Value::String(s) => Some(s.to_string()),
            _ => None,
        })
    };
    if type_tag.as_deref() == Some("Date") {
        let receiver = Value::Object(obj.clone());
        let prefer = if hint == "string" { ["toString", "valueOf"] } else { ["valueOf", "toString"] };
        for m in &prefer {
            let r = dispatch(ctx, &receiver, m, &[]);
            if !matches!(r, Value::Object(_) | Value::Undefined) {
                return r;
            }
        }
    }
    let methods: &[&str] = if hint == "string" {
        &["toString", "valueOf"]
    } else {
        &["valueOf", "toString"]
    };
    // Route through `dispatch` (which mirrors the JS method-call
    // protocol — sets `__js_this`, then calls). Direct ctx.invoke
    // bypasses __js_this binding, so the user's body sees a stale
    // global and reads `.v` on null/undefined → throws TypeError.
    let receiver = Value::Object(obj.clone());
    for m in methods {
        let exists = lookup_method_via_proto(&obj, m).is_some();
        if !exists { continue; }
        let result = dispatch(ctx, &receiver, m, &[]);
        if !matches!(result, Value::Object(_) | Value::Undefined) {
            return result;
        }
    }
    // Class instances with a `__type` tag get the spec-shaped
    // `[object <Name>]` rather than `[object]` (the Vybe Display
    // default for Ordinary).
    let tag = {
        let o = obj.lock().unwrap();
        o.properties.get("__type").map(|t| format!("{}", t))
    };
    match tag {
        Some(t) if !t.is_empty() => Value::String(Arc::from(format!("[object {}]", t).as_str())),
        _ => Value::String(Arc::from("[object Object]")),
    }
}


/// If `arg` is a RegExp object (Object stamped with `__type=RegExp`),
/// extract its `(source, flags)` strings. Otherwise return None so the
/// caller falls back to literal-string handling. Mirrors the shape
/// produced by `ecma:regexp.new`.
fn regex_pattern(arg: Option<&Value>) -> Option<(String, String)> {
    let Some(Value::Object(obj)) = arg else { return None; };
    let o = obj.lock().unwrap();
    let type_tag = o.properties.get("__type")?;
    if !matches!(type_tag, Value::String(s) if s.as_ref() == "RegExp") {
        return None;
    }
    let src = match o.properties.get("source")? {
        Value::String(s) => s.to_string(),
        other => format!("{}", other),
    };
    let flags = match o.properties.get("flags") {
        Some(Value::String(s)) => s.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    };
    Some((src, flags))
}

// ── WeakMap dynamic dispatch (for when type isn't known at compile time) ─────

fn dispatch_weakmap(obj: Arc<Mutex<Object>>, method: &str, args: &[Value]) -> Value {
    match method {
        "get" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(key, Value::Object(_)) { return Value::Undefined; }
            let m = obj.lock().unwrap();
            if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                let ko = keys_obj.lock().unwrap();
                if let ObjectKind::Array(ref keys) = ko.kind {
                    if let Some(pos) = wm_key_ptr_find(keys, &key) {
                        drop(ko);
                        if let ObjectKind::Array(ref values) = m.kind {
                            return values.get(pos).cloned().unwrap_or(Value::Undefined);
                        }
                    }
                }
            }
            Value::Undefined
        }
        "set" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            let val = args.get(1).cloned().unwrap_or(Value::Undefined);
            if !matches!(key, Value::Object(_)) { return Value::Object(obj); }
            // find or insert
            let existing = {
                let m = obj.lock().unwrap();
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref keys) = ko.kind {
                        wm_key_ptr_find(keys, &key)
                    } else { None }
                } else { None }
            };
            let mut m = obj.lock().unwrap();
            if let Some(pos) = existing {
                if let ObjectKind::Array(ref mut values) = m.kind { values[pos] = val; }
            } else {
                if let ObjectKind::Array(ref mut values) = m.kind { values.push(val); }
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP).cloned() {
                    let mut ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref mut keys) = ko.kind { keys.push(key); }
                }
            }
            drop(m);
            Value::Object(obj)
        }
        "has" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(key, Value::Object(_)) { return Value::Bool(false); }
            let m = obj.lock().unwrap();
            if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                let ko = keys_obj.lock().unwrap();
                if let ObjectKind::Array(ref keys) = ko.kind {
                    return Value::Bool(wm_key_ptr_find(keys, &key).is_some());
                }
            }
            Value::Bool(false)
        }
        "delete" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(key, Value::Object(_)) { return Value::Bool(false); }
            let mut m = obj.lock().unwrap();
            let pos = if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                let ko = keys_obj.lock().unwrap();
                if let ObjectKind::Array(ref keys) = ko.kind { wm_key_ptr_find(keys, &key) } else { None }
            } else { None };
            if let Some(pos) = pos {
                if let ObjectKind::Array(ref mut values) = m.kind { values.remove(pos); }
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP).cloned() {
                    let mut ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref mut keys) = ko.kind { keys.remove(pos); }
                }
                return Value::Bool(true);
            }
            Value::Bool(false)
        }
        _ => Value::Undefined,
    }
}

// ── WeakSet dynamic dispatch ──────────────────────────────────────────────────

fn dispatch_weakset(obj: Arc<Mutex<Object>>, method: &str, args: &[Value]) -> Value {
    match method {
        "add" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(v, Value::Object(_)) { return Value::Object(obj); }
            let mut so = obj.lock().unwrap();
            let already = if let ObjectKind::Array(ref vs) = so.kind {
                wm_key_ptr_find(vs, &v).is_some()
            } else { false };
            if !already {
                if let ObjectKind::Array(ref mut vs) = so.kind { vs.push(v); }
            }
            drop(so);
            Value::Object(obj)
        }
        "has" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(v, Value::Object(_)) { return Value::Bool(false); }
            let so = obj.lock().unwrap();
            if let ObjectKind::Array(ref vs) = so.kind {
                return Value::Bool(wm_key_ptr_find(vs, &v).is_some());
            }
            Value::Bool(false)
        }
        "delete" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(v, Value::Object(_)) { return Value::Bool(false); }
            let mut so = obj.lock().unwrap();
            let pos = if let ObjectKind::Array(ref vs) = so.kind {
                wm_key_ptr_find(vs, &v)
            } else { None };
            if let Some(pos) = pos {
                if let ObjectKind::Array(ref mut vs) = so.kind { vs.remove(pos); }
                return Value::Bool(true);
            }
            Value::Bool(false)
        }
        _ => Value::Undefined,
    }
}
