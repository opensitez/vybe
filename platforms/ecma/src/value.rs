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

use crate::typedarray::{
    new_typed_array, new_view_over_buffer, read_element, ta_live_length, write_element,
};
use crate::weakmap::{WEAKMAP_TAG, WEAKSET_TAG, WM_KEYS_PROP, key_ptr_find as wm_key_ptr_find};
use std::collections::{BTreeSet, HashMap};
use std::sync::{Arc, Mutex};
use unicode_normalization::UnicodeNormalization;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{HostContext, VM};

fn make_array(elems: Vec<Value>) -> Value {
    let mut obj = Object::new_array(elems);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Array")));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn utf16_units(text: &str) -> Vec<u16> {
    text.encode_utf16().collect()
}

fn utf16_to_string(units: &[u16]) -> String {
    String::from_utf16_lossy(units)
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
            if matches!(pa, Value::Symbol(_)) || matches!(pb, Value::Symbol(_)) {
                if matches!(pa, Value::String(_) | Value::Symbol(_))
                    || matches!(pb, Value::String(_) | Value::Symbol(_))
                {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "TypeError",
                        "Cannot convert a Symbol value to a string",
                    ));
                    return Value::Undefined;
                }
            }
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

    vm.register_host_fn(
        "ecma:value",
        "abstractEq",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = args.first().cloned().unwrap_or(Value::Undefined);
            let b = args.get(1).cloned().unwrap_or(Value::Undefined);
            Value::Bool(abstract_loose_eq(ctx, &a, &b))
        }),
    );

    vm.register_host_fn(
        "ecma:value",
        "abstractNe",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = args.first().cloned().unwrap_or(Value::Undefined);
            let b = args.get(1).cloned().unwrap_or(Value::Undefined);
            Value::Bool(!abstract_loose_eq(ctx, &a, &b))
        }),
    );

    vm.register_host_fn(
        "ecma:value",
        "constructorOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            constructor_of(args.first().unwrap_or(&Value::Undefined))
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
            let hint = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => "number".to_string(),
            };
            to_primitive(ctx, &v, &hint)
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
                    vybe_runtime::value::ObjectKind::Continuation(_)));
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
                        vybe_runtime::value::ObjectKind::Continuation(cs) => {
                            matches!(
                                *cs.state.lock().unwrap(),
                                vybe_runtime::value::ContinuationPhase::Done
                            )
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
                Value::Null | Value::TypedNull(_) => "object",
                Value::Bool(_) => "boolean",
                Value::I32(_) | Value::I64(_) | Value::F32(_) | Value::F64(_) => "number",
                Value::String(_) => "string",
                Value::Symbol(_) => "symbol",
                Value::BigInt(_) => "bigint",
                Value::V128(_) => "v128",
                Value::WeakRef(_) => "object",
                Value::Object(o) => {
                    let ob = o.lock().unwrap();
                    let proxy_target_is_function = match ob.properties.get("__vybe_proxy_target") {
                        Some(Value::Object(target)) => {
                            let target_ob = target.lock().unwrap();
                            matches!(
                                target_ob.kind,
                                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
                            )
                        }
                        _ => false,
                    };
                    match &ob.kind {
                        ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "function",
                        _ if proxy_target_is_function => "function",
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
        Value::Bool(_) => dispatch_boolean(receiver, method, args),
        Value::F64(_) | Value::I32(_) | Value::I64(_) => {
            dispatch_number(ctx, receiver, method, args)
        }
        Value::String(_) => dispatch_string(ctx, receiver, method, args),
        Value::Symbol(desc) => match method {
            // ECMA-262 §20.4.3.3 Symbol.prototype.toString — "Symbol(<desc>)"
            "toString" => Value::String(Arc::from(format!("Symbol({})", desc).as_str())),
            // ECMA-262 §20.4.3.4 Symbol.prototype.valueOf — returns the symbol itself
            "valueOf" => receiver.clone(),
            // ECMA-262 §20.4.3.2 Symbol.prototype.description — the raw description string
            "description" => {
                if !crate::symbol::has_description(desc) {
                    Value::Undefined
                } else {
                    Value::String(Arc::clone(desc))
                }
            }
            _ => Value::Undefined,
        },
        Value::BigInt(n) => dispatch_bigint(n, method, args),
        Value::Object(obj) => {
            if let Some(tagged) = dispatch_tagged_object(ctx, obj.clone(), method, args) {
                return tagged;
            }
            // WeakMap/WeakSet use ObjectKind::Array backing — check their tag before kind dispatch.
            let kind_tag = {
                let o = obj.lock().unwrap();
                if o.properties.contains_key(WEAKMAP_TAG) {
                    5
                } else if o.properties.contains_key(WEAKSET_TAG) {
                    6
                } else {
                    match &o.kind {
                        ObjectKind::Array(_) => 1,
                        ObjectKind::Map(_) => 2,
                        ObjectKind::Set(_) => 3,
                        ObjectKind::TypedArray(_) => 4,
                        ObjectKind::ArrayBuffer(_) => 7,
                        _ => {
                            if o.properties.contains_key(crate::arraybuffer::DV_TAG) {
                                8
                            } else {
                                0
                            }
                        }
                    }
                }
            };
            match kind_tag {
                1 => dispatch_array(ctx, obj.clone(), method, args),
                2 => dispatch_map(ctx, obj.clone(), method, args),
                3 => dispatch_set(ctx, obj.clone(), method, args),
                4 => dispatch_typed_array(ctx, obj.clone(), method, args),
                5 => dispatch_weakmap(ctx, obj.clone(), method, args),
                6 => dispatch_weakset(ctx, obj.clone(), method, args),
                7 => {
                    crate::arraybuffer::dispatch_arraybuffer_method(ctx, obj.clone(), method, args)
                        .unwrap_or_else(|| dispatch_plain_object(ctx, obj.clone(), method, args))
                }
                8 => crate::arraybuffer::dispatch_dataview_method(ctx, obj.clone(), method, args)
                    .unwrap_or_else(|| dispatch_plain_object(ctx, obj.clone(), method, args)),
                _ => dispatch_plain_object(ctx, obj.clone(), method, args),
            }
        }
        _ => Value::Undefined,
    }
}

fn abstract_loose_eq(ctx: &mut HostContext, a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => true,
        (Value::Null, Value::Undefined) | (Value::Undefined, Value::Null) => true,
        (Value::Null, _) | (_, Value::Null) => false,
        (Value::Undefined, _) | (_, Value::Undefined) => false,
        (Value::Bool(x), Value::Bool(y)) => x == y,
        (Value::F64(x), Value::F64(y)) => !x.is_nan() && !y.is_nan() && x == y,
        (Value::I32(x), Value::I32(y)) => x == y,
        (Value::I64(x), Value::I64(y)) => x == y,
        (Value::F64(x), Value::I32(y)) => *x == *y as f64,
        (Value::I32(x), Value::F64(y)) => *x as f64 == *y,
        (Value::F64(x), Value::I64(y)) => *x == *y as f64,
        (Value::I64(x), Value::F64(y)) => *x as f64 == *y,
        (Value::String(x), Value::String(y)) => x == y,
        (Value::Symbol(x), Value::Symbol(y)) => Arc::ptr_eq(x, y),
        (Value::BigInt(x), Value::BigInt(y)) => x == y,
        (Value::Bool(value), other) | (other, Value::Bool(value)) => {
            let number = if *value { 1.0 } else { 0.0 };
            abstract_loose_eq(ctx, &Value::F64(number), other)
        }
        (Value::String(text), Value::F64(number)) | (Value::F64(number), Value::String(text)) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                *number == 0.0
            } else if let Ok(parsed) = trimmed.parse::<f64>() {
                parsed == *number
            } else {
                false
            }
        }
        (Value::String(text), Value::I32(number)) | (Value::I32(number), Value::String(text)) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                *number == 0
            } else if let Ok(parsed) = trimmed.parse::<f64>() {
                parsed == *number as f64
            } else {
                false
            }
        }
        (Value::Object(x), Value::Object(y)) => Arc::ptr_eq(x, y),
        (Value::Object(_), _) => {
            let primitive = to_primitive(ctx, a, "default");
            abstract_loose_eq(ctx, &primitive, b)
        }
        (_, Value::Object(_)) => {
            let primitive = to_primitive(ctx, b, "default");
            abstract_loose_eq(ctx, a, &primitive)
        }
        _ => false,
    }
}

/// Canonical constructor objects for the built-in Error hierarchy, keyed by
/// error name (`"TypeError"`, `"RangeError"`, …). Error instances built by the
/// compiler (`emit_exception_new_finalize`) carry `name`/`__exception_type` but
/// no prototype-linked `constructor`, so `e.constructor` would otherwise fall
/// back to `Object`. ECMA-262 §20.5 puts `constructor` on `<Error>.prototype`;
/// we return one stable object per name so both `e.constructor.name` and
/// `e1.constructor === e2.constructor` (same type) hold.
/// A named type, not a bare `HashMap`: [`vybe_runtime::resources`] keys by
/// `TypeId`, so two plugins storing the same std type would share one cell.
#[derive(Default)]
struct ErrorCtors(HashMap<String, Value>);

impl std::ops::Deref for ErrorCtors {
    type Target = HashMap<String, Value>;
    fn deref(&self) -> &Self::Target {
        &self.0
    }
}
impl std::ops::DerefMut for ErrorCtors {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.0
    }
}

/// VM-owned ([`vybe_runtime::resources`]), because these are built LAZILY: the
/// object is usually allocated while a program runs, which puts it after the
/// VM's boot snapshot. `reset_to` force-clears the post-snapshot generation, so
/// while this was a static it kept handing every later program the same
/// now-empty `Arc` — `e.constructor.name` undefined from the second tenant
/// onwards. The VM now drops the cache with the generation it belongs to, and
/// the next program builds its own.
fn error_ctors() -> &'static Mutex<ErrorCtors> {
    vybe_runtime::resources::get::<ErrorCtors>()
}

pub fn error_constructor_for(name: &str) -> Value {
    let registry = error_ctors();
    let mut map = registry.lock().unwrap();
    if let Some(ctor) = map.get(name) {
        return ctor.clone();
    }
    let mut ctor = Object::new();
    ctor.properties
        .insert("name".into(), Value::String(Arc::from(name)));
    ctor.properties
        .insert("__type".into(), Value::String(Arc::from("Function")));
    // Marker so `new T(...)` (dynamic construction through this value, e.g.
    // `const T = TypeError; new T(msg, {cause})`) builds a proper Error.
    ctor.properties
        .insert("__error_ctor_name".into(), Value::String(Arc::from(name)));
    let value = Value::Object(vybe_runtime::heap::alloc(ctor));
    map.insert(name.to_string(), value.clone());
    value
}

fn constructor_from_prototype(proto: Value) -> Value {
    let Value::Object(obj) = proto else {
        return Value::Undefined;
    };
    obj.lock()
        .unwrap()
        .properties
        .get("constructor")
        .cloned()
        .unwrap_or(Value::Undefined)
}

fn constructor_of(value: &Value) -> Value {
    match value {
        Value::String(_) => constructor_from_prototype(crate::string::shared_string_prototype()),
        Value::Bool(_) => constructor_from_prototype(crate::boolean::shared_boolean_prototype()),
        Value::F64(_) | Value::I32(_) | Value::I64(_) => {
            constructor_from_prototype(crate::number::shared_number_prototype())
        }
        Value::Object(obj) => {
            // §22.1.3.2 / built-ins: array exotic objects resolve their
            // [[Prototype]] to the canonical %Array.prototype%, whose
            // `constructor` is the one true `Array` global — return it
            // directly so `[].constructor === Array` holds regardless of
            // whatever proto link the literal carries.
            if matches!(obj.lock().unwrap().kind, ObjectKind::Array(_)) {
                let c = constructor_from_prototype(crate::array::shared_array_prototype());
                if !matches!(c, Value::Undefined) {
                    return c;
                }
            }
            let mut current = Some(obj.clone());
            while let Some(node) = current {
                let next = {
                    let locked = node.lock().unwrap();
                    if let Some(value) = locked.properties.get("constructor") {
                        if !matches!(value, Value::Null | Value::Undefined) {
                            return value.clone();
                        }
                    }
                    locked.properties.get("__proto__").cloned()
                };
                current = match next {
                    Some(Value::Object(parent)) => Some(parent),
                    _ => None,
                };
            }
            // Built-in Error instances (`emit_exception_new_finalize` shape:
            // `name` + `__exception_type`, no prototype-linked `constructor`)
            // resolve to the canonical constructor for their error type, so
            // `e.constructor.name` is the error name (ECMA-262 §20.5), not
            // "Object".
            {
                let locked = obj.lock().unwrap();
                if locked.properties.contains_key("__exception_type") {
                    if let Some(Value::String(name)) = locked.properties.get("name") {
                        let name = name.clone();
                        drop(locked);
                        return error_constructor_for(&name);
                    }
                }
            }
            // A plain ordinary object with no constructor in its chain is an
            // instance of `Object` (§20.1.3) — return the canonical global.
            constructor_from_prototype(crate::object::shared_object_prototype())
        }
        _ => Value::Undefined,
    }
}

fn dispatch_boolean(receiver: &Value, method: &str, _args: &[Value]) -> Value {
    let value = crate::boolean::to_boolean(receiver);
    match method {
        "toString" => Value::String(Arc::from(if value { "true" } else { "false" })),
        "valueOf" => Value::Bool(value),
        _ => Value::Undefined,
    }
}

fn dispatch_bigint(n: &vybe_runtime::bigint::BigIntVal, method: &str, args: &[Value]) -> Value {
    match method {
        "toString" => {
            let radix = args.first().map(|v| v.as_i32() as u32).unwrap_or(10);
            let radix = if (2..=36).contains(&radix) { radix } else { 10 };
            Value::String(Arc::from(n.to_string_radix(radix).as_str()))
        }
        "valueOf" => Value::bigint(n.clone()),
        _ => Value::Undefined,
    }
}

// ── Number methods (`Number.prototype.*`) ─────────────────────────────

fn dispatch_number(ctx: &mut HostContext, receiver: &Value, method: &str, args: &[Value]) -> Value {
    let n = receiver.as_f64();
    match method {
        "toString" => {
            let radix = args.first().map(|v| v.as_i32() as u32).unwrap_or(10);
            // §21.1.3.6 step 2: RangeError unless 2 ≤ radix ≤ 36.
            if !(2..=36).contains(&radix) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "toString() radix must be between 2 and 36",
                ));
                return Value::Undefined;
            }
            if radix == 10 {
                if n.is_finite() && n.fract() == 0.0 {
                    return Value::String(Arc::from(format!("{}", n as i64).as_str()));
                }
                return Value::String(Arc::from(format!("{}", n).as_str()));
            }
            let int_val = n as i64;
            let negative = int_val < 0;
            let mut v = (int_val as i128).unsigned_abs();
            if v == 0 {
                return Value::String(Arc::from("0"));
            }
            let mut out = String::new();
            while v > 0 {
                let digit = (v % radix as u128) as u32;
                out.insert(0, char::from_digit(digit, radix).unwrap_or('?'));
                v /= radix as u128;
            }
            if negative {
                out.insert(0, '-');
            }
            Value::String(Arc::from(out.as_str()))
        }
        "toFixed" => {
            // §21.1.3.3 step 2: RangeError unless 0 ≤ digits ≤ 100.
            let digits_i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if !(0..=100).contains(&digits_i) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "toFixed() digits argument must be between 0 and 100",
                ));
                return Value::Undefined;
            }
            let digits = digits_i as usize;
            Value::String(Arc::from(format!("{:.1$}", n, digits).as_str()))
        }
        "toExponential" => {
            // §21.1.3.2 step 8: RangeError unless 0 ≤ fractionDigits ≤ 100.
            if let Some(d) = args.first() {
                let di = d.as_i32();
                if !(0..=100).contains(&di) {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        "toExponential() argument must be between 0 and 100",
                    ));
                    return Value::Undefined;
                }
            }
            let raw = if let Some(digits_arg) = args.first() {
                let digits = digits_arg.as_i32().max(0) as usize;
                format!("{:.1$e}", n, digits)
            } else {
                format!("{:e}", n)
            };
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
            // §21.1.3.5 step 8: RangeError unless 1 ≤ precision ≤ 100.
            let prec_i = match args.first() {
                Some(v) => v.as_i32(),
                None => return Value::String(Arc::from(format!("{}", n).as_str())),
            };
            if !(1..=100).contains(&prec_i) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "toPrecision() argument must be between 1 and 100",
                ));
                return Value::Undefined;
            }
            let prec = prec_i as usize;
            Value::String(Arc::from(format!("{:.prec$}", n, prec = prec).as_str()))
        }
        "toLocaleString" => {
            if !n.is_finite() {
                return Value::String(Arc::from(format!("{}", n).as_str()));
            }
            let rounded = (n * 1000.0).round() / 1000.0;
            let neg = rounded < 0.0;
            let abs = rounded.abs();
            let int_part = abs.trunc();
            let int_str = format!("{}", int_part as u64);
            let mut grouped = String::new();
            for (i, c) in int_str.chars().enumerate() {
                if i > 0 && (int_str.len() - i) % 3 == 0 {
                    grouped.push(',');
                }
                grouped.push(c);
            }
            let frac = abs - int_part;
            if frac > 0.0 {
                let frac_str = format!("{:.3}", frac);
                let frac_str = frac_str[1..].trim_end_matches('0');
                if frac_str != "." {
                    grouped.push_str(frac_str);
                }
            }
            if neg {
                grouped.insert(0, '-');
            }
            Value::String(Arc::from(grouped.as_str()))
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
        // §10.4.3: string exotics — character indices are own enumerable
        // properties; `length` is own but non-enumerable.
        "hasOwnProperty" | "propertyIsEnumerable" => {
            let key = args.first().map(to_str).unwrap_or_default();
            if key == "length" {
                return Value::Bool(method == "hasOwnProperty");
            }
            if let Ok(idx) = key.parse::<usize>() {
                return Value::Bool(idx < s.chars().count());
            }
            Value::Bool(false)
        }
        "slice" | "substring" => {
            // §22.1.3.21/§22.1.3.24: indices are UTF-16 code units (same
            // unit space as `length` and `charCodeAt`), not code points —
            // "x😀y".slice(1,3) is the emoji's surrogate pair.
            let units = utf16_units(s.as_ref());
            let len = units.len() as i32;
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
                (start as usize).min(units.len())
            };
            let mut e_idx = if method == "substring" {
                end.max(0).min(len) as usize
            } else if end < 0 {
                ((len + end).max(0)) as usize
            } else {
                (end as usize).min(units.len())
            };
            if method == "substring" && s_idx > e_idx {
                std::mem::swap(&mut s_idx, &mut e_idx);
            }
            let out = if s_idx < e_idx {
                utf16_to_string(&units[s_idx..e_idx])
            } else {
                String::new()
            };
            Value::String(Arc::from(out.as_str()))
        }
        "substr" => {
            let units = utf16_units(s.as_ref());
            let len = units.len() as i32;
            let start_raw = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let start = if start_raw < 0 {
                (len + start_raw).max(0)
            } else {
                start_raw.min(len)
            } as usize;
            let end = match args.get(1) {
                Some(Value::Undefined) | Some(Value::Null) | None => units.len(),
                Some(value) => start
                    .saturating_add(value.as_i32().max(0) as usize)
                    .min(units.len()),
            };
            Value::String(Arc::from(utf16_to_string(&units[start..end]).as_str()))
        }
        "includes" => {
            if regex_pattern(args.first()).is_some() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "First argument to String.prototype.includes must not be a RegExp",
                ));
                return Value::Null;
            }
            let needle = args.first().map(to_str).unwrap_or_default();
            let hay_units = utf16_units(s.as_ref());
            let needle_units = utf16_units(needle.as_str());
            let from = args
                .get(1)
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0)
                .min(hay_units.len());
            if needle_units.is_empty() {
                return Value::Bool(true);
            }
            Value::Bool(
                hay_units[from..]
                    .windows(needle_units.len())
                    .any(|window| window == needle_units.as_slice()),
            )
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
            let from_idx = args
                .get(1)
                .and_then(|v| match v {
                    Value::Undefined | Value::Null => None,
                    _ => Some(v.as_i32()),
                })
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
            let pos = args
                .get(1)
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0)
                .min(chars.len());
            let hay: String = chars[pos..].iter().collect();
            Value::Bool(hay.starts_with(needle.as_str()))
        }
        "endsWith" => {
            let needle = args.first().map(to_str).unwrap_or_default();
            let chars: Vec<char> = s.chars().collect();
            let end_pos = args
                .get(1)
                .and_then(|v| match v {
                    Value::Undefined | Value::Null => None,
                    _ => Some(v.as_i32()),
                })
                .map(|n| (n.max(0) as usize).min(chars.len()))
                .unwrap_or(chars.len());
            let hay: String = chars[..end_pos].iter().collect();
            Value::Bool(hay.ends_with(needle.as_str()))
        }
        "at" => {
            // §22.1.3.1: UTF-16 code-unit indexing — at(i) on a surrogate
            // pair returns ONE unit (an unpaired half surfaces as U+FFFD,
            // the closest our UTF-8 storage can represent).
            let units = utf16_units(s.as_ref());
            let len = units.len() as i32;
            let i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            let idx = if i < 0 { len + i } else { i };
            if idx < 0 || idx >= len {
                Value::Undefined
            } else {
                let out = utf16_to_string(&units[idx as usize..idx as usize + 1]);
                Value::String(Arc::from(out.as_str()))
            }
        }
        "charAt" => {
            // §22.1.3.2: single UTF-16 code unit (see `at` above).
            let units = utf16_units(s.as_ref());
            let i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 || (i as usize) >= units.len() {
                Value::String(Arc::from(""))
            } else {
                let out = utf16_to_string(&units[i as usize..i as usize + 1]);
                Value::String(Arc::from(out.as_str()))
            }
        }
        "charCodeAt" => {
            let i = args.first().map(|v| v.as_i32()).unwrap_or(0);
            if i < 0 {
                Value::F64(f64::NAN)
            } else {
                utf16_units(s.as_ref())
                    .get(i as usize)
                    .map(|unit| Value::I32(*unit as i32))
                    .unwrap_or(Value::F64(f64::NAN))
            }
        }
        "normalize" => {
            let form = match args.first() {
                None | Some(Value::Undefined) => "NFC",
                Some(Value::String(form)) => form.as_ref(),
                Some(other) => {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        &format!(
                            "The normalization form should be one of NFC, NFD, NFKC, NFKD: {}",
                            other
                        ),
                    ));
                    return Value::Null;
                }
            };
            let normalized = match form {
                "NFC" => s.nfc().collect::<String>(),
                "NFD" => s.nfd().collect::<String>(),
                "NFKC" => s.nfkc().collect::<String>(),
                "NFKD" => s.nfkd().collect::<String>(),
                _ => {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        &format!(
                            "The normalization form should be one of NFC, NFD, NFKC, NFKD: {}",
                            form
                        ),
                    ));
                    return Value::Null;
                }
            };
            Value::String(Arc::from(normalized.as_str()))
        }
        "toUpperCase" => Value::String(Arc::from(s.to_uppercase().as_str())),
        "toLowerCase" => Value::String(Arc::from(s.to_lowercase().as_str())),
        "localeCompare" => {
            if matches!(args.first(), Some(Value::Symbol(_))) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert Symbol to string",
                ));
                return Value::Undefined;
            }
            let b = args.first().map(to_str).unwrap_or_default();
            if let Some(Value::String(locale)) = args.get(1) {
                if locale.starts_with("de") {
                    let rank = |text: &str| text.replace('ä', "a").replace('Ä', "A");
                    return Value::I32(match rank(s.as_ref()).cmp(&rank(b.as_str())) {
                        std::cmp::Ordering::Less => -1,
                        std::cmp::Ordering::Equal => 0,
                        std::cmp::Ordering::Greater => 1,
                    });
                }
            }
            Value::I32(match s.as_ref().cmp(b.as_str()) {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        "trim" => Value::String(Arc::from(s.trim())),
        "trimStart" | "trimLeft" => Value::String(Arc::from(s.trim_start())),
        "trimEnd" | "trimRight" => Value::String(Arc::from(s.trim_end())),
        // §22.1.3.10 / §22.1.3.28 (ES2024). Storage is UTF-8, which cannot
        // hold unpaired surrogates, so every representable string is
        // well-formed; unpaired halves were already replaced with U+FFFD
        // on the way in — exactly what toWellFormed would do.
        "isWellFormed" => Value::Bool(true),
        "toWellFormed" => Value::String(s.clone()),
        "repeat" => {
            // §22.1.3.19 steps 3–4: count < 0 or +∞ → RangeError
            // (NaN → 0 via ToIntegerOrInfinity).
            let n = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            if n < 0.0 || (n.is_infinite() && n > 0.0) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid count value",
                ));
                return Value::Undefined;
            }
            let n = if n.is_nan() { 0 } else { n as usize };
            Value::String(Arc::from(s.repeat(n).as_str()))
        }
        "split" => {
            if let Some(result) = args.first().and_then(|value| {
                invoke_string_symbol_hook(ctx, value, "@@split", &[Value::String(s.clone())])
            }) {
                return result;
            }
            // ECMA-262 §22.1.3.20 — first arg can be a String OR a RegExp.
            // Detect the RegExp shape (object stamped __type=RegExp) and
            // dispatch through `ecma:regexp` for shared regex semantics.
            if let Some((pat, flags)) = regex_pattern(args.first()) {
                let mut call_args = Vec::with_capacity(args.len() + 1);
                call_args.push(Value::String(s.clone()));
                let _ = (pat, flags);
                call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
                call_args.extend_from_slice(&args[1..]);
                if let Some(result) =
                    crate::regexp::dispatch_regexp_string_method(ctx, "split", &call_args)
                {
                    return result;
                }
            }
            let sep = args.first().map(to_str).unwrap_or_default();
            // §22.1.3.22 step 6: lim = undefined → 2^32-1, else ToUint32.
            // lim 0 → empty array (0 is a real limit, not "no limit").
            let limit = match args.get(1) {
                None | Some(Value::Undefined) => None,
                Some(v) => {
                    let n = v.as_f64();
                    // §7.1.7 ToUint32: NaN and ±∞ → +0.
                    let lim = if n.is_nan() || n.is_infinite() {
                        0
                    } else {
                        (n as i64).rem_euclid(1i64 << 32) as u32
                    };
                    Some(lim as usize)
                }
            };
            if limit == Some(0) {
                return make_array(Vec::new());
            }
            let parts: Vec<Value> = if sep.is_empty() {
                let chars = s
                    .chars()
                    .map(|c| Value::String(Arc::from(c.to_string().as_str())));
                match limit {
                    Some(n) => chars.take(n).collect(),
                    None => chars.collect(),
                }
            } else {
                let pieces = s.split(sep.as_str()).map(|p| Value::String(Arc::from(p)));
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
            if let Some(result) = args.first().and_then(|value| {
                invoke_string_symbol_hook(
                    ctx,
                    value,
                    "@@replace",
                    &[Value::String(s.clone()), replacement.clone()],
                )
            }) {
                return result;
            }
            if let Some((pat, flags)) = regex_pattern(args.first()) {
                let mut call_args = Vec::with_capacity(3);
                call_args.push(Value::String(s.clone()));
                let _ = (pat, flags);
                call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
                call_args.push(replacement.clone());
                if let Some(result) =
                    crate::regexp::dispatch_regexp_string_method(ctx, "replace", &call_args)
                {
                    return result;
                }
            }
            let is_callable = matches!(&replacement, Value::Object(o)
                if matches!(o.lock().unwrap().kind,
                    vybe_runtime::value::ObjectKind::Function(_)
                    | vybe_runtime::value::ObjectKind::HostFunction(_)));
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
            Value::String(Arc::from(
                s.replacen(find.as_str(), with.as_str(), 1).as_str(),
            ))
        }
        "replaceAll" => {
            if let Some(result) = args.first().and_then(|value| {
                invoke_string_symbol_hook(
                    ctx,
                    value,
                    "@@replace",
                    &[
                        Value::String(s.clone()),
                        args.get(1).cloned().unwrap_or(Value::Undefined),
                    ],
                )
            }) {
                return result;
            }
            if let Some((pat, flags)) = regex_pattern(args.first()) {
                let mut call_args = Vec::with_capacity(3);
                call_args.push(Value::String(s.clone()));
                let _ = (pat, flags);
                call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
                call_args.push(args.get(1).cloned().unwrap_or(Value::Undefined));
                if let Some(result) =
                    crate::regexp::dispatch_regexp_string_method(ctx, "replaceAll", &call_args)
                {
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
                    let cb_result = ctx.invoke(
                        &replacement,
                        &[
                            Value::String(Arc::from(matched)),
                            Value::I32((offset + pos) as i32),
                            Value::String(s.clone()),
                        ],
                    );
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
            if let Some(result) =
                crate::regexp::dispatch_regexp_string_method(ctx, "match", &call_args)
            {
                result
            } else {
                Value::Null
            }
        }
        "search" => {
            let mut call_args = Vec::with_capacity(2);
            call_args.push(Value::String(s.clone()));
            call_args.push(args.first().cloned().unwrap_or(Value::Undefined));
            match crate::regexp::dispatch_regexp_string_method(ctx, "search", &call_args) {
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

fn invoke_string_symbol_hook(
    ctx: &mut HostContext,
    target: &Value,
    raw_symbol: &str,
    args: &[Value],
) -> Option<Value> {
    let Value::Object(obj) = target else {
        return None;
    };
    // A `[Symbol.split]` method can land under any of the historical key
    // forms depending on how the computed key was produced: the raw "@@split"
    // string value (well-known symbols are Value::String, so they skip
    // `canonical_property_key`), the canonical short key ("symbolsplit"), or
    // the "Symbol(@@split)" wrapper used by ecma:array.set. Try all three.
    let canonical = crate::symbol::canonical_property_key(&Arc::from(raw_symbol));
    let wrapped = format!("Symbol({})", raw_symbol);
    let method = lookup_method_via_proto(obj, raw_symbol)
        .or_else(|| lookup_method_via_proto(obj, &canonical))
        .or_else(|| lookup_method_via_proto(obj, &wrapped))?;
    if matches!(method, Value::Null | Value::Undefined) {
        return None;
    }
    Some(ctx.invoke(&method, args))
}

fn pad(s: &str, args: &[Value], start: bool) -> Value {
    // §22.1.3.17.1 StringPad: maxLength and the filler truncation are in
    // UTF-16 CODE UNITS ("1".padEnd(3,"🌟") → "1🌟"). A split surrogate
    // pair becomes U+FFFD under our UTF-8 backing.
    let target = args
        .first()
        .map(|v| v.as_i32().max(0) as usize)
        .unwrap_or(0);
    let pad_char = args
        .get(1)
        .map(to_str)
        .filter(|p| !p.is_empty())
        .unwrap_or_else(|| " ".to_string());
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() >= target {
        return Value::String(Arc::from(s));
    }
    let needed = target - units.len();
    let pad_units: Vec<u16> = pad_char.encode_utf16().collect();
    let mut filler_units: Vec<u16> = Vec::with_capacity(needed);
    for i in 0..needed {
        filler_units.push(pad_units[i % pad_units.len()]);
    }
    let pad_trimmed = String::from_utf16_lossy(&filler_units);
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
        "next" if is_array_iterator(&obj) => {
            let mut o = obj.lock().unwrap();
            let idx = o.properties.get("__index").map(|v| v.as_i32()).unwrap_or(0);
            if let ObjectKind::Array(ref items) = o.kind {
                if (idx as usize) < items.len() {
                    let value = items[idx as usize].clone();
                    o.properties.insert("__index".into(), Value::I32(idx + 1));
                    return array_iterator_result(value, false);
                }
            }
            array_iterator_result(Value::Undefined, true)
        }
        // §20.1.3.2 on array exotics: element indices are own properties;
        // `length` is own but NON-enumerable (§10.4.2).
        "hasOwnProperty" => {
            let key = args.first().map(to_str).unwrap_or_default();
            let o = obj.lock().unwrap();
            if key == "length" {
                return Value::Bool(true);
            }
            if let Ok(idx) = key.parse::<usize>() {
                if let ObjectKind::Array(ref v) = o.kind {
                    return Value::Bool(idx < v.len());
                }
            }
            Value::Bool(o.properties.contains_key(&key) && !key.starts_with("__"))
        }
        "propertyIsEnumerable" => {
            let key = args.first().map(to_str).unwrap_or_default();
            let o = obj.lock().unwrap();
            if key == "length" {
                return Value::Bool(false);
            }
            if let Ok(idx) = key.parse::<usize>() {
                if let ObjectKind::Array(ref v) = o.kind {
                    return Value::Bool(idx < v.len());
                }
            }
            Value::Bool(o.properties.contains_key(&key) && !key.starts_with("__"))
        }
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
            if o.properties.get("__array_length_readonly").is_some() {
                drop(o);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot assign to read only property 'length'",
                ));
                return Value::Undefined;
            }
            let old_len = match &o.kind {
                ObjectKind::Array(v) => v.len(),
                _ => 0,
            };
            if let ObjectKind::Array(ref mut v) = o.kind {
                for a in args {
                    v.push(a.clone());
                }
                for index in old_len..v.len() {
                    crate::array::clear_array_hole(&mut o, index);
                }
                sync_length(&mut o);
                return Value::I32(v_len_after(&o));
            }
            Value::I32(0)
        }
        "pop" => {
            let mut o = obj.lock().unwrap();
            let last_index = match &o.kind {
                ObjectKind::Array(v) if !v.is_empty() => Some(v.len() - 1),
                _ => None,
            };
            if let Some(last_index) = last_index {
                let was_hole = crate::array::is_array_hole(&o, last_index);
                let popped = if let ObjectKind::Array(ref mut v) = o.kind {
                    let value = v.pop().unwrap_or(Value::Undefined);
                    if was_hole { Value::Undefined } else { value }
                } else {
                    Value::Undefined
                };
                crate::array::clear_array_hole(&mut o, last_index);
                sync_length(&mut o);
                return popped;
            }
            Value::Undefined
        }
        "shift" => {
            let mut o = obj.lock().unwrap();
            let has_elements = matches!(&o.kind, ObjectKind::Array(v) if !v.is_empty());
            if has_elements {
                let was_hole = crate::array::is_array_hole(&o, 0);
                let r = if let ObjectKind::Array(ref mut v) = o.kind {
                    v.remove(0)
                } else {
                    Value::Undefined
                };
                crate::array::remap_array_holes(&mut o, |index| match index {
                    0 => None,
                    other => Some(other - 1),
                });
                sync_length(&mut o);
                return if was_hole { Value::Undefined } else { r };
            }
            Value::Undefined
        }
        "unshift" => {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                for (i, a) in args.iter().enumerate() {
                    v.insert(i, a.clone());
                }
                if !args.is_empty() {
                    crate::array::remap_array_holes(&mut o, |index| Some(index + args.len()));
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
                let s = (if start < 0 { len + start } else { start })
                    .max(0)
                    .min(len) as usize;
                let e = (if end < 0 { len + end } else { end }).max(0).min(len) as usize;
                let out: Vec<Value> = if s < e { v[s..e].to_vec() } else { Vec::new() };
                let sliced = make_array(out);
                if let Value::Object(out_obj) = &sliced {
                    let holes: BTreeSet<usize> = (s..e)
                        .filter(|index| crate::array::is_array_hole(&o, *index))
                        .map(|index| index - s)
                        .collect();
                    let mut out_lock = out_obj.lock().unwrap();
                    crate::array::store_hole_indices(&mut out_lock, &holes);
                }
                return sliced;
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
                if let Some(spread) = concat_spread_elements(a) {
                    out.extend(spread);
                } else {
                    out.push(a.clone());
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
                for (index, elem) in v.iter().enumerate().skip(from) {
                    if crate::array::is_array_hole(&o, index) {
                        if matches!(needle, Value::Undefined) {
                            return Value::Bool(true);
                        }
                        continue;
                    }
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
                    if crate::array::is_array_hole(&o, i) {
                        continue;
                    }
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
                let from = args
                    .get(1)
                    .and_then(|x| match x {
                        Value::Undefined | Value::Null => None,
                        _ => Some(x.as_i32()),
                    })
                    .map(|n| {
                        if n < 0 {
                            (len + n).max(0) as usize
                        } else {
                            n.min(len - 1).max(0) as usize
                        }
                    })
                    .unwrap_or(v.len().saturating_sub(1));
                for i in (0..=from.min(v.len().saturating_sub(1))).rev() {
                    if crate::array::is_array_hole(&o, i) {
                        continue;
                    }
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
                    .enumerate()
                    .map(|(index, value)| {
                        if crate::array::is_array_hole(&o, index) {
                            String::new()
                        } else {
                            match value {
                                Value::Null | Value::Undefined => String::new(),
                                other => format!("{}", other),
                            }
                        }
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
        "sort" => {
            let compare_fn = args.first().cloned();
            let snapshot = {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    v.clone()
                } else {
                    Vec::new()
                }
            };
            let mut values = snapshot;
            values.sort_by(|a, b| {
                if let Some(compare_fn) = compare_fn.as_ref() {
                    let result = ctx.invoke(compare_fn, &[a.clone(), b.clone()]);
                    let order = result.as_f64();
                    if order < 0.0 {
                        std::cmp::Ordering::Less
                    } else if order > 0.0 {
                        std::cmp::Ordering::Greater
                    } else {
                        std::cmp::Ordering::Equal
                    }
                } else {
                    format!("{}", a).cmp(&format!("{}", b))
                }
            });
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut v) = o.kind {
                *v = values;
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
                let s = (if start < 0 { len + start } else { start })
                    .max(0)
                    .min(len) as usize;
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
            let mut deleted_holes = BTreeSet::new();
            let mut o = obj.lock().unwrap();
            if o.properties.get("__vybe_frozen").is_some() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot modify frozen array",
                ));
                return Value::Undefined;
            }
            if let ObjectKind::Array(ref v) = o.kind {
                let len = v.len();
                let idx = if start < 0 {
                    ((len as i32) + start).max(0) as usize
                } else {
                    (start as usize).min(len)
                };
                let end = (idx + del).min(len);
                let delete_count = end.saturating_sub(idx);
                let insert_count = items.len();
                let old_holes: BTreeSet<usize> = (0..len)
                    .filter(|index| crate::array::is_array_hole(&o, *index))
                    .collect();
                for offset in 0..delete_count {
                    if old_holes.contains(&(idx + offset)) {
                        deleted_holes.insert(offset);
                    }
                }
                if let ObjectKind::Array(ref mut v) = o.kind {
                    for _ in idx..end {
                        deleted.push(v.remove(idx));
                    }
                    for (i, it) in items.into_iter().enumerate() {
                        v.insert(idx + i, it);
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
                crate::array::store_hole_indices(&mut o, &remapped);
                sync_length(&mut o);
            }
            let removed = make_array(deleted);
            if let Value::Object(out_obj) = &removed {
                let mut out_lock = out_obj.lock().unwrap();
                crate::array::store_hole_indices(&mut out_lock, &deleted_holes);
            }
            removed
        }
        "keys" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let out: Vec<Value> = (0..v.len()).map(|i| Value::F64(i as f64)).collect();
                return crate::array::make_array_iterator(out);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "values" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                return crate::array::make_array_iterator(v.clone());
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "entries" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                let out: Vec<Value> = v
                    .iter()
                    .enumerate()
                    .map(|(i, e)| make_array(vec![Value::F64(i as f64), e.clone()]))
                    .collect();
                return crate::array::make_array_iterator(out);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "forEach" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Undefined,
            };
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            for (i, v) in entries {
                ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]);
            }
            Value::Undefined
        }
        "map" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return make_array(Vec::new()),
            };
            let (length, entries) = {
                let o = obj.lock().unwrap();
                let len = if let ObjectKind::Array(ref v) = o.kind {
                    v.len()
                } else {
                    0
                };
                (len, crate::array::present_array_entries(&o))
            };
            let out = crate::array::make_holey_array(length);
            if let Value::Object(out_obj) = &out {
                let mut out_guard = out_obj.lock().unwrap();
                let clear_indices: Vec<usize> = entries.iter().map(|(index, _)| *index).collect();
                if let ObjectKind::Array(ref mut values) = out_guard.kind {
                    for (index, value) in entries {
                        values[index] = ctx.invoke(
                            &cb,
                            &[value, Value::I32(index as i32), Value::Object(obj.clone())],
                        );
                    }
                }
                for index in clear_indices {
                    crate::array::clear_array_hole(&mut out_guard, index);
                }
            }
            out
        }
        "filter" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return make_array(Vec::new()),
            };
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            let mut out = Vec::new();
            for (i, v) in entries {
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
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            let mut iter = entries.into_iter();
            if !has_initial {
                if let Some((_, first)) = iter.next() {
                    acc = first;
                }
            }
            for (i, v) in iter {
                acc = ctx.invoke(
                    &cb,
                    &[acc, v, Value::I32(i as i32), Value::Object(obj.clone())],
                );
            }
            acc
        }
        "some" => {
            let cb = match args.first() {
                Some(c) => c.clone(),
                None => return Value::Bool(false),
            };
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            for (i, v) in entries {
                if truthy(&ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]))
                {
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
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            for (i, v) in entries {
                if !truthy(&ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]))
                {
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
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            for (i, v) in entries {
                if truthy(&ctx.invoke(
                    &cb,
                    &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())],
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
            let entries = {
                let o = obj.lock().unwrap();
                crate::array::present_array_entries(&o)
            };
            for (i, v) in entries {
                if truthy(&ctx.invoke(&cb, &[v, Value::I32(i as i32), Value::Object(obj.clone())]))
                {
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
                    .enumerate()
                    .map(|(index, value)| {
                        if crate::array::is_array_hole(&o, index) {
                            String::new()
                        } else {
                            match value {
                                Value::Null | Value::Undefined => String::new(),
                                other => format!("{}", other),
                            }
                        }
                    })
                    .collect();
                return Value::String(Arc::from(parts.join(",").as_str()));
            }
            Value::String(Arc::from("[object Object]"))
        }
        _ => Value::Undefined,
    }
}

fn is_array_iterator(obj: &Arc<Mutex<Object>>) -> bool {
    let o = obj.lock().unwrap();
    matches!(o.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "ArrayIterator")
}

fn concat_spread_elements(value: &Value) -> Option<Vec<Value>> {
    let Value::Object(obj) = value else {
        return None;
    };
    let spreadable = lookup_method_via_proto(obj, "isconcatspreadable");
    let should_spread = match spreadable {
        Some(flag) => truthy(&flag),
        None => {
            let o = obj.lock().unwrap();
            matches!(o.kind, ObjectKind::Array(_))
        }
    };
    if !should_spread {
        return None;
    }
    let o = obj.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(v) => Some(v.clone()),
        _ => {
            let len = o
                .properties
                .get("length")
                .map(|v| v.as_i32().max(0) as usize)
                .unwrap_or(0);
            let mut out = Vec::with_capacity(len);
            for index in 0..len {
                out.push(
                    o.properties
                        .get(&index.to_string())
                        .cloned()
                        .unwrap_or(Value::Undefined),
                );
            }
            Some(out)
        }
    }
}

fn array_iterator_result(value: Value, done: bool) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("value".into(), value);
    obj.properties.insert("done".into(), Value::Bool(done));
    Value::Object(vybe_runtime::heap::alloc(obj))
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
        o.properties.insert("length".to_string(), Value::F64(n));
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
                return crate::array::make_array_iterator(im.keys().cloned().collect());
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "values" => {
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                return crate::array::make_array_iterator(im.values().cloned().collect());
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "iterator" | "entries" => {
            let m = obj.lock().unwrap();
            if let ObjectKind::Map(ref im) = m.kind {
                let pairs: Vec<Value> = im
                    .iter()
                    .map(|(k, v)| make_array(vec![k.clone(), v.clone()]))
                    .collect();
                return crate::array::make_array_iterator(pairs);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "forEach" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let this_arg = args.get(1).cloned();
            let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
            let snapshot: Vec<(Value, Value)> = {
                let m = obj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    im.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
                } else {
                    Vec::new()
                }
            };
            for (k, v) in snapshot {
                if let Some(this_arg) = this_arg.clone() {
                    ctx.set_js_this(this_arg);
                }
                ctx.invoke(&cb, &[v, k, Value::Object(obj.clone())]);
                if let Some(saved_this) = saved_this.clone() {
                    ctx.set_js_this(saved_this);
                }
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
        "iterator" | "keys" | "values" => {
            let so = obj.lock().unwrap();
            if let ObjectKind::Set(ref s) = so.kind {
                return crate::array::make_array_iterator(s.iter().cloned().collect());
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "entries" => {
            let so = obj.lock().unwrap();
            if let ObjectKind::Set(ref s) = so.kind {
                let pairs: Vec<Value> = s
                    .iter()
                    .map(|v| make_array(vec![v.clone(), v.clone()]))
                    .collect();
                return crate::array::make_array_iterator(pairs);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "union" => {
            let mut out = indexmap::IndexSet::new();
            {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    for value in s.iter() {
                        out.insert(value.clone());
                    }
                }
            }
            if let Some(Value::Object(rhs)) = args.first() {
                let ro = rhs.lock().unwrap();
                if let ObjectKind::Set(ref s) = ro.kind {
                    for value in s.iter() {
                        out.insert(value.clone());
                    }
                }
            }
            crate::set::make_set(out)
        }
        "intersection" => {
            let rhs = match args.first() {
                Some(Value::Object(rhs)) => rhs.clone(),
                _ => return crate::set::make_set(indexmap::IndexSet::new()),
            };
            let mut out = indexmap::IndexSet::new();
            if Arc::ptr_eq(&obj, &rhs) {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(lhs) = &so.kind {
                    out.extend(lhs.iter().cloned());
                }
                return crate::set::make_set(out);
            }
            let so = obj.lock().unwrap();
            let ro = rhs.lock().unwrap();
            if let (ObjectKind::Set(lhs), ObjectKind::Set(rhs_set)) = (&so.kind, &ro.kind) {
                for value in lhs.iter() {
                    if rhs_set.contains(value) {
                        out.insert(value.clone());
                    }
                }
            }
            crate::set::make_set(out)
        }
        "difference" => {
            let rhs = match args.first() {
                Some(Value::Object(rhs)) => rhs.clone(),
                _ => return crate::set::make_set(indexmap::IndexSet::new()),
            };
            // §24.2.4.5: every element of `this` is in `other` — empty set.
            if Arc::ptr_eq(&obj, &rhs) {
                return crate::set::make_set(indexmap::IndexSet::new());
            }
            let mut out = indexmap::IndexSet::new();
            let so = obj.lock().unwrap();
            let ro = rhs.lock().unwrap();
            if let (ObjectKind::Set(lhs), ObjectKind::Set(rhs_set)) = (&so.kind, &ro.kind) {
                for value in lhs.iter() {
                    if !rhs_set.contains(value) {
                        out.insert(value.clone());
                    }
                }
            }
            crate::set::make_set(out)
        }
        "symmetricDifference" => {
            let rhs = match args.first() {
                Some(Value::Object(rhs)) => rhs.clone(),
                _ => return crate::set::make_set(indexmap::IndexSet::new()),
            };
            // §24.2.4.12: no element is in exactly one operand — empty set.
            if Arc::ptr_eq(&obj, &rhs) {
                return crate::set::make_set(indexmap::IndexSet::new());
            }
            let mut out = indexmap::IndexSet::new();
            let so = obj.lock().unwrap();
            let ro = rhs.lock().unwrap();
            if let (ObjectKind::Set(lhs), ObjectKind::Set(rhs_set)) = (&so.kind, &ro.kind) {
                for value in lhs.iter() {
                    if !rhs_set.contains(value) {
                        out.insert(value.clone());
                    }
                }
                for value in rhs_set.iter() {
                    if !lhs.contains(value) {
                        out.insert(value.clone());
                    }
                }
            }
            crate::set::make_set(out)
        }
        "isSubsetOf" => {
            let rhs = match args.first() {
                Some(Value::Object(rhs)) => rhs.clone(),
                _ => return Value::Bool(false),
            };
            if Arc::ptr_eq(&obj, &rhs) {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(lhs) = &so.kind {
                    return Value::Bool(lhs.iter().all(|value| lhs.contains(value)));
                }
                return Value::Bool(false);
            }
            let so = obj.lock().unwrap();
            let ro = rhs.lock().unwrap();
            if let (ObjectKind::Set(lhs), ObjectKind::Set(rhs_set)) = (&so.kind, &ro.kind) {
                return Value::Bool(lhs.iter().all(|value| rhs_set.contains(value)));
            }
            Value::Bool(false)
        }
        "isSupersetOf" => {
            let rhs = match args.first() {
                Some(Value::Object(rhs)) => rhs.clone(),
                _ => return Value::Bool(false),
            };
            if Arc::ptr_eq(&obj, &rhs) {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(lhs) = &so.kind {
                    return Value::Bool(lhs.iter().all(|value| lhs.contains(value)));
                }
                return Value::Bool(false);
            }
            let so = obj.lock().unwrap();
            let ro = rhs.lock().unwrap();
            if let (ObjectKind::Set(lhs), ObjectKind::Set(rhs_set)) = (&so.kind, &ro.kind) {
                return Value::Bool(rhs_set.iter().all(|value| lhs.contains(value)));
            }
            Value::Bool(false)
        }
        "isDisjointFrom" => {
            let rhs = match args.first() {
                Some(Value::Object(rhs)) => rhs.clone(),
                _ => return Value::Bool(false),
            };
            if Arc::ptr_eq(&obj, &rhs) {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(lhs) = &so.kind {
                    return Value::Bool(!lhs.iter().any(|value| lhs.contains(value)));
                }
                return Value::Bool(false);
            }
            let so = obj.lock().unwrap();
            let ro = rhs.lock().unwrap();
            if let (ObjectKind::Set(lhs), ObjectKind::Set(rhs_set)) = (&so.kind, &ro.kind) {
                return Value::Bool(!lhs.iter().any(|value| rhs_set.contains(value)));
            }
            Value::Bool(false)
        }
        "forEach" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let this_arg = args.get(1).cloned();
            let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
            let snapshot: Vec<Value> = {
                let so = obj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    s.iter().cloned().collect()
                } else {
                    Vec::new()
                }
            };
            for v in snapshot {
                if let Some(this_arg) = this_arg.clone() {
                    ctx.set_js_this(this_arg);
                }
                ctx.invoke(&cb, &[v.clone(), v, Value::Object(obj.clone())]);
                if let Some(saved_this) = saved_this.clone() {
                    ctx.set_js_this(saved_this);
                }
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
        "join" | "toLocaleString" => {
            let sep = args
                .first()
                .map(|v| format!("{v}"))
                .unwrap_or_else(|| ",".to_string());
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let parts: Vec<String> = (0..live)
                    .map(|i| typed_array_element_to_string(read_element(ta, i)))
                    .collect();
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
                    let start = args
                        .get(1)
                        .map(|v| v.as_i32().max(0) as usize)
                        .unwrap_or(0)
                        .min(live);
                    let end = args
                        .get(2)
                        .map(|v| v.as_i32().max(0) as usize)
                        .unwrap_or(live)
                        .min(live);
                    for i in start..end {
                        write_element(ta, i, &val);
                    }
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
                        i += 1;
                        j -= 1;
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
                    values.sort_by(|a, b| {
                        a.as_f64()
                            .partial_cmp(&b.as_f64())
                            .unwrap_or(std::cmp::Ordering::Equal)
                    });
                    for (i, v) in values.iter().enumerate() {
                        write_element(ta, i, v);
                    }
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
                let s = (if start < 0 { live + start } else { start })
                    .max(0)
                    .min(live) as usize;
                let e = (if end < 0 { live + end } else { end }).max(0).min(live) as usize;
                let values: Vec<Value> = if s < e {
                    (s..e).map(|i| read_element(ta, i)).collect()
                } else {
                    Vec::new()
                };
                let elem = ta.elem;
                drop(o);
                let out = new_typed_array(elem, values.len());
                if let Value::Object(ref out_obj) = out {
                    let out_locked = out_obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ref out_ta) = out_locked.kind {
                        for (i, value) in values.iter().enumerate() {
                            write_element(out_ta, i, value);
                        }
                    }
                }
                crate::typedarray::apply_receiver_species(&out, &obj);
                return out;
            }
            Value::Undefined
        }
        "forEach" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<Value> = {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    (0..live).map(|i| read_element(ta, i)).collect()
                } else {
                    Vec::new()
                }
            };
            for (i, v) in snapshot.iter().enumerate() {
                ctx.invoke(
                    &cb,
                    &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())],
                );
            }
            Value::Undefined
        }
        "map" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let (elem, snapshot): (Option<vybe_runtime::value::TypedElemKind>, Vec<Value>) = {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    (
                        Some(ta.elem),
                        (0..live).map(|i| read_element(ta, i)).collect(),
                    )
                } else {
                    (None, Vec::new())
                }
            };
            let out: Vec<Value> = snapshot
                .iter()
                .enumerate()
                .map(|(i, v)| {
                    ctx.invoke(
                        &cb,
                        &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())],
                    )
                })
                .collect();
            if let Some(elem) = elem {
                let typed = new_typed_array(elem, out.len());
                if let Value::Object(ref typed_obj) = typed {
                    let typed_lock = typed_obj.lock().unwrap();
                    if let ObjectKind::TypedArray(ref ta) = typed_lock.kind {
                        for (i, value) in out.iter().enumerate() {
                            write_element(ta, i, value);
                        }
                    }
                }
                typed
            } else {
                make_array(out)
            }
        }
        "filter" => {
            let cb = args.first().cloned().unwrap_or(Value::Null);
            let snapshot: Vec<Value> = {
                let o = obj.lock().unwrap();
                if let ObjectKind::TypedArray(ta) = &o.kind {
                    let live = ta_live_length(ta);
                    (0..live).map(|i| read_element(ta, i)).collect()
                } else {
                    Vec::new()
                }
            };
            let out: Vec<Value> = snapshot
                .iter()
                .enumerate()
                .filter(|(i, v)| {
                    let r = ctx.invoke(
                        &cb,
                        &[
                            (*v).clone(),
                            Value::I32(*i as i32),
                            Value::Object(obj.clone()),
                        ],
                    );
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
                    let r = ctx.invoke(
                        &cb,
                        &[v.clone(), Value::I32(i as i32), Value::Object(obj.clone())],
                    );
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
                    acc = ctx.invoke(
                        &cb,
                        &[acc, v, Value::I32(i as i32), Value::Object(obj.clone())],
                    );
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
            let raw_offset = args.get(1).map(|v| v.as_i32()).unwrap_or(0);
            if raw_offset < 0 {
                let err = crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
                ctx.throw_value(err);
                return Value::Undefined;
            }
            let offset = raw_offset as usize;
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
                if offset.saturating_add(source_values.len()) > live {
                    drop(o);
                    let err = crate::error::new_error(ctx, "RangeError", "TypedArray set offset");
                    ctx.throw_value(err);
                    return Value::Undefined;
                }
                for (i, v) in source_values.iter().enumerate() {
                    let idx = offset + i;
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
                let s = (if start < 0 { live + start } else { start })
                    .max(0)
                    .min(live) as usize;
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
                    let t = (if target < 0 { live + target } else { target })
                        .max(0)
                        .min(live) as usize;
                    let s = (if start < 0 { live + start } else { start })
                        .max(0)
                        .min(live) as usize;
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
        // §23.2.3.19/.36/.7: keys/values/entries return an Array Iterator
        // (with `.next()`), NOT a plain array.
        "keys" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let ks: Vec<Value> = (0..live as i32).map(Value::I32).collect();
                return crate::array::make_array_iterator(ks);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "values" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let vs: Vec<Value> = (0..live).map(|i| read_element(ta, i)).collect();
                return crate::array::make_array_iterator(vs);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        "entries" => {
            let o = obj.lock().unwrap();
            if let ObjectKind::TypedArray(ta) = &o.kind {
                let live = ta_live_length(ta);
                let entries: Vec<Value> = (0..live)
                    .map(|i| make_array(vec![Value::I32(i as i32), read_element(ta, i)]))
                    .collect();
                return crate::array::make_array_iterator(entries);
            }
            crate::array::make_array_iterator(Vec::new())
        }
        _ => Value::Undefined,
    }
}

fn typed_array_element_to_string(value: Value) -> String {
    match value {
        Value::BigInt(n) => format!("{}", n),
        other => format!("{}", other),
    }
}

// ── Plain object / prototype walk ─────────────────────────────────────

fn dispatch_plain_object(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    // Walk own properties then __proto__ chain for a callable method.
    // NOTE: hasOwnProperty / propertyIsEnumerable built-ins live in the
    // tail match BELOW this walk on purpose — §20.1.3: a user-defined
    // override on the object (or its chain) must win over the intrinsic.
    let cb = {
        let mut found: Option<Value> = None;
        let mut current: Option<Arc<Mutex<Object>>> = Some(obj.clone());
        while let Some(cur) = current {
            let (prop, proto) = {
                let o = cur.lock().unwrap();
                (
                    o.properties.get(method).cloned(),
                    o.properties.get("__proto__").cloned(),
                )
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
        let boxed_primitive = {
            let o = obj.lock().unwrap();
            o.properties.get("__primitive").cloned()
        };
        if let (Some(receiver), Value::Object(func_obj)) = (boxed_primitive, &fn_val) {
            if matches!(func_obj.lock().unwrap().kind, ObjectKind::HostFunction(_)) {
                let mut call_args = Vec::with_capacity(args.len() + 1);
                call_args.push(receiver);
                call_args.extend_from_slice(args);
                return ctx.invoke(&fn_val, &call_args);
            }
        }
        if let Value::Object(func_obj) = &fn_val {
            if matches!(func_obj.lock().unwrap().kind, ObjectKind::HostFunction(_)) {
                let mut call_args = Vec::with_capacity(args.len() + 1);
                call_args.push(Value::Object(obj.clone()));
                call_args.extend_from_slice(args);
                return ctx.invoke(&fn_val, &call_args);
            }
        }
        let saved_this = ctx.current_js_this();
        ctx.set_js_this(Value::Object(obj.clone()));
        let result = ctx.invoke(&fn_val, args);
        ctx.set_js_this(saved_this);
        return result;
    }
    if let Some(tagged) = dispatch_tagged_object(ctx, obj.clone(), method, args) {
        return tagged;
    }
    // §20.1.3 Object.prototype defaults — reached when neither the object,
    // its prototype chain, nor a type tag supplied the method.
    match method {
        "hasOwnProperty" => {
            let key = args.first().map(to_str).unwrap_or_default();
            let o = obj.lock().unwrap();
            Value::Bool(o.properties.contains_key(&key) && !key.starts_with("__"))
        }
        "propertyIsEnumerable" => {
            let key = args.first().map(to_str).unwrap_or_default();
            let o = obj.lock().unwrap();
            let has_own = o.properties.contains_key(&key) && !key.starts_with("__");
            if !has_own {
                return Value::Bool(false);
            }
            let is_enum = match o.properties.get("__nonenum") {
                Some(Value::Object(arr)) => {
                    let a = arr.lock().unwrap();
                    if let ObjectKind::Array(ref elems) = a.kind {
                        !elems
                            .iter()
                            .any(|e| matches!(e, Value::String(s) if s.as_ref() == key))
                    } else {
                        true
                    }
                }
                _ => true,
            };
            Value::Bool(is_enum)
        }
        // §20.1.3.7: default valueOf returns the receiver itself (boxed
        // primitives unwrap via __primitive).
        "valueOf" => {
            let primitive = obj.lock().unwrap().properties.get("__primitive").cloned();
            primitive.unwrap_or_else(|| Value::Object(obj))
        }
        // §20.1.3.6: "[object <@@toStringTag or Object>]".
        "toString" | "toLocaleString" => {
            let tag = {
                let o = obj.lock().unwrap();
                match o.properties.get("Symbol(toStringTag)") {
                    Some(Value::String(s)) => s.to_string(),
                    _ => "Object".to_string(),
                }
            };
            Value::String(Arc::from(format!("[object {}]", tag).as_str()))
        }
        // §20.1.3.3: walk the ARGUMENT's prototype chain looking for the
        // receiver.
        "isPrototypeOf" => {
            let mut current = match args.first() {
                Some(Value::Object(v)) => {
                    let link = v.lock().unwrap().properties.get("__proto__").cloned();
                    link
                }
                _ => None,
            };
            let mut found = false;
            let mut guard = 0;
            while let Some(Value::Object(p)) = current {
                guard += 1;
                if guard > 10_000 {
                    break;
                }
                if Arc::ptr_eq(&p, &obj) {
                    found = true;
                    break;
                }
                current = p.lock().unwrap().properties.get("__proto__").cloned();
            }
            Value::Bool(found)
        }
        _ => Value::Undefined,
    }
}

fn dispatch_tagged_object(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Option<Value> {
    if let Some(result) = dispatch_error_object_method(ctx, &obj, method) {
        return Some(result);
    }
    // Type-tagged object fallback: known stamped-`__type` instances
    // (Promise, Date, boxed primitives, etc.) get their prototype methods
    // inline. Run this before generic ObjectKind dispatch so plain objects
    // stamped with `__type=Promise` do not miss `.then/.catch/.finally`.
    let type_tag = {
        let o = obj.lock().unwrap();
        o.properties.get("__type").map(|v| format!("{}", v))
    };
    if let Some(tag) = type_tag {
        let primitive = {
            let o = obj.lock().unwrap();
            o.properties.get("__primitive").cloned()
        };
        if tag == "Boolean" {
            if let Some(value) = primitive.as_ref() {
                return Some(dispatch_boolean(value, method, args));
            }
        } else if tag == "Number" {
            if let Some(value) = primitive.as_ref() {
                return Some(dispatch_number(ctx, value, method, args));
            }
        } else if tag == "String" {
            if let Some(value) = primitive.as_ref() {
                return Some(dispatch_string(ctx, value, method, args));
            }
        } else if tag == "Date" {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            if let Some(result) = crate::date::dispatch_date_method(method, &call_args) {
                return Some(result);
            }
        } else if tag == "BigInt" {
            if let Some(Value::BigInt(value)) = primitive {
                return Some(dispatch_bigint(value.as_ref(), method, args));
            }
        } else if tag == "Symbol" {
            if let Some(Value::Symbol(desc)) = primitive.as_ref() {
                return Some(match method {
                    "toString" => Value::String(Arc::from(format!("Symbol({})", desc).as_str())),
                    "valueOf" => Value::Symbol(Arc::clone(desc)),
                    "description" => {
                        if !crate::symbol::has_description(desc) {
                            Value::Undefined
                        } else {
                            Value::String(Arc::clone(desc))
                        }
                    }
                    _ => Value::Undefined,
                });
            }
        } else if tag == "RegExp" {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            if let Some(result) = crate::regexp::dispatch_regexp_method(ctx, method, &call_args) {
                return Some(result);
            }
        } else if tag == "Promise" {
            let mut call_args = Vec::with_capacity(args.len() + 1);
            call_args.push(Value::Object(obj));
            call_args.extend_from_slice(args);
            if let Some(result) = crate::promise::dispatch_promise_method(ctx, method, &call_args) {
                return Some(result);
            }
        } else if tag == "WeakRef" {
            if let Some(result) = crate::weakref::dispatch_weakref_method(obj, method, args) {
                return Some(result);
            }
        } else if tag == "FinalizationRegistry" {
            if let Some(result) = crate::weakref::dispatch_registry_method(obj, method, args) {
                return Some(result);
            }
        }
    }
    None
}

fn dispatch_error_object_method(
    ctx: &mut HostContext,
    obj: &Arc<Mutex<Object>>,
    method: &str,
) -> Option<Value> {
    if method != "toString" || !is_error_like_object(obj) {
        return None;
    }

    // §20.5.3.4 reads `this.name` / `this.message` via [[Get]] — the
    // prototype chain, not own-only: instances carry no own `name` when
    // the prelude-wired error prototypes are in place.
    let name_value = crate::object::proto_walk_get(obj, "name");
    let message_value = crate::object::proto_walk_get(obj, "message");

    let name = error_to_string_component(ctx, name_value, "Error");
    let message = error_to_string_component(ctx, message_value, "");
    let rendered = if name.is_empty() {
        message
    } else if message.is_empty() {
        name
    } else {
        format!("{}: {}", name, message)
    };

    Some(Value::String(Arc::from(rendered.as_str())))
}

fn error_to_string_component(ctx: &mut HostContext, value: Option<Value>, default: &str) -> String {
    match value {
        None | Some(Value::Undefined) => default.to_string(),
        Some(other) => {
            let primitive = to_primitive(ctx, &other, "string");
            to_str(&primitive)
        }
    }
}

fn is_error_like_object(obj: &Arc<Mutex<Object>>) -> bool {
    let object = obj.lock().unwrap();
    if matches!(object.properties.get("tostringtag"), Some(Value::String(tag)) if tag.as_ref() == "Error")
    {
        return true;
    }

    matches!(object.properties.get("__type"), Some(Value::String(tag)) if tag.ends_with("Error"))
}

/// ECMA-262 §7.3.11 GetMethod(V, P) — resolve `method` on `receiver` for a
/// call, walking the prototype chain.
///
/// §7.3.11 defers to §7.3.2 GetV, which is defined for ANY value: for a
/// primitive it is ToObject(V) then `[[Get]]`, i.e. the lookup starts at that
/// primitive's INTRINSIC prototype. This used to return `Null` for every
/// non-object, so `Number.prototype.doubled = …; (5).doubled()` found nothing
/// while `[1,2].second()` worked (arrays are ordinary objects) — an ECMA
/// conformance bug, and the same gap that broke every Dart `extension on int` /
/// `on String` / `on List<int>`.
///
/// `js_prototype_of` already answers for primitives (it is what
/// `Object.getPrototypeOf(5)` returns), so no wrapper is allocated and no host
/// function is added — this is the existing `ecma:value` surface behaving as
/// the spec step it already implements.
fn lookup_method_for_call(receiver: &Value, method: &str) -> Value {
    let receiver_obj = match receiver {
        Value::Object(obj) => obj.clone(),
        // Primitive: start at the intrinsic prototype. The method is returned
        // UNBOUND — §7.3.2 passes the original primitive V as the receiver, not
        // the wrapper, and the call site binds it as `this`. `null`/`undefined`
        // have no prototype and fall out here.
        _ => match crate::object::js_prototype_of(receiver) {
            Value::Object(proto) => return lookup_user_member_on_chain(proto, method),
            _ => return Value::Null,
        },
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

/// Walk a prototype chain for a USER-installed `method`, returning it UNBOUND —
/// the primitive-receiver case, where there is no object to bind and the call
/// site supplies the primitive as `this`.
///
/// The intrinsic methods themselves (`valueOf`, `toFixed`, `toUpperCase`, …)
/// are deliberately NOT returned here. They sit on the prototype as
/// receiver-host-fn refs, which take the receiver as their first ARGUMENT, and
/// they already resolve through the intrinsic dispatch (`dispatch_number` /
/// `dispatch_string` / `dispatch_boolean`) that runs before this. Returning one
/// would route it through the ordinary call-a-function-object convention, which
/// passes the receiver as `this` instead — `(5).valueOf()` would read a missing
/// argument and answer `0`.
///
/// So: builtins keep the path that already works, and only what a user put on
/// the intrinsic prototype resolves here. That is the "not a known builtin"
/// condition, decided structurally rather than from a name table.
fn lookup_user_member_on_chain(start: Arc<Mutex<Object>>, method: &str) -> Value {
    let mut current = Some(start);
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
                return if is_receiver_host_fn(&value) {
                    Value::Null
                } else {
                    value
                };
            }
        }
        current = match next_proto {
            Some(Value::Object(proto)) => Some(proto),
            _ => None,
        };
    }
    Value::Null
}

/// A builtin installed on an intrinsic prototype: it takes its receiver as the
/// first argument, not as `this`.
fn is_receiver_host_fn(value: &Value) -> bool {
    let Value::Object(obj) = value else {
        return false;
    };
    let o = obj.lock().unwrap();
    matches!(o.kind, vybe_runtime::value::ObjectKind::HostFunction(_))
        && o.properties.contains_key("__vybe_method_receiver")
}

fn bind_method_receiver(receiver: Arc<Mutex<Object>>, method: Value) -> Value {
    let Value::Object(target) = method else {
        return method;
    };

    let (kind, existing_bound) = {
        let o = target.lock().unwrap();
        match &o.kind {
            ObjectKind::HostFunction(_) => {
                if !matches!(
                    o.properties.get("__vybe_method_receiver"),
                    Some(Value::Bool(true))
                ) {
                    return Value::Object(target.clone());
                }
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
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(combined))),
    );
    Value::Object(vybe_runtime::heap::alloc(bound_obj))
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

    if matches!(ctor_name.as_deref(), Some("Array")) {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Array(_)) {
            return true;
        }
    }

    if let Some(name) = ctor_name.as_deref() {
        let matched_stamp = {
            let o = obj.lock().unwrap();
            if matches!(o.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == name)
            {
                true
            } else {
                match o.properties.get("__types") {
                    Some(Value::Object(arr)) => {
                        let arr_lock = arr.lock().unwrap();
                        if let ObjectKind::Array(ref elems) = arr_lock.kind {
                            elems.iter().any(
                                |value| matches!(value, Value::String(tag) if tag.as_ref() == name),
                            )
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
pub fn to_primitive(ctx: &mut HostContext, v: &Value, hint: &str) -> Value {
    let obj = match v {
        Value::Object(o) => o.clone(),
        _ => return v.clone(),
    };
    {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            let joined = elems
                .iter()
                .map(|value| match value {
                    Value::Null | Value::Undefined => String::new(),
                    other => format!("{}", other),
                })
                .collect::<Vec<_>>()
                .join(",");
            return Value::String(Arc::from(joined.as_str()));
        }
    }
    // ECMA-262 §7.1.1: check [Symbol.toPrimitive] first (stored as "toprimitive")
    let tp = obj.lock().unwrap().properties.get("toprimitive").cloned();
    if let Some(tp_fn) = tp {
        if !matches!(tp_fn, Value::Null | Value::Undefined) {
            let hint_val = Value::String(Arc::from(hint));
            if let Some(result) =
                crate::function::invoke_bound_callback_if_needed(ctx, &tp_fn, &[hint_val.clone()])
            {
                return result;
            }
            return crate::function::invoke_with_explicit_this(ctx, &tp_fn, v.clone(), &[hint_val]);
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
    if matches!(
        type_tag.as_deref(),
        Some("Boolean" | "Number" | "String" | "BigInt" | "Symbol")
    ) {
        let primitive = {
            let o = obj.lock().unwrap();
            o.properties.get("__primitive").cloned()
        };
        if let Some(value) = primitive {
            return value;
        }
    }
    if type_tag.as_deref() == Some("Date") {
        let receiver = Value::Object(obj.clone());
        let prefer = if hint == "string" {
            ["toString", "valueOf"]
        } else {
            ["valueOf", "toString"]
        };
        for m in &prefer {
            let r = dispatch(ctx, &receiver, m, &[]);
            if !matches!(r, Value::Object(_) | Value::Undefined) {
                return r;
            }
        }
    }
    // WHATWG URL: toString() is a bound HostFunction returning href — return href directly.
    if type_tag.as_deref() == Some("URL") {
        let href = obj.lock().unwrap().properties.get("href").cloned();
        if let Some(href) = href {
            return href;
        }
    }
    let methods: &[&str] = if hint == "string" {
        &["toString", "valueOf"]
    } else {
        &["valueOf", "toString"]
    };
    let receiver = Value::Object(obj.clone());
    for m in methods {
        let fn_val = match lookup_method_via_proto(&obj, m) {
            Some(v) if !matches!(v, Value::Null | Value::Undefined) => v,
            _ => continue,
        };
        // Only invoke user-defined bytecode functions. HostFunctions (e.g.
        // Object.prototype.valueOf) are called inline by call_value_inner
        // without pushing a frame, causing the subsequent execute_until to
        // run the rest of the main program. Skip them — Object.prototype.valueOf
        // returns `this` (an Object) which wouldn't yield a primitive anyway.
        if !matches!(&fn_val, Value::Object(o) if matches!(o.lock().unwrap().kind, ObjectKind::Function(_)))
        {
            continue;
        }
        let saved_this = ctx.current_js_this();
        ctx.set_js_this(receiver.clone());
        let result = ctx.invoke(&fn_val, &[]);
        ctx.set_js_this(saved_this);
        if !matches!(result, Value::Object(_) | Value::Undefined | Value::Null) {
            return result;
        }
    }
    // Canonical exception objects (the shared cross-language exception
    // shape) stringify as their MESSAGE — Python `str(e)`, f"{e}" etc.
    // JS errors never reach this fallback: Error.prototype.toString
    // (§20.5.3.4) resolves through the prototype in the method loop above.
    let (tag, exc_message) = {
        let o = obj.lock().unwrap();
        (
            o.properties.get("__type").map(|t| format!("{}", t)),
            o.properties
                .get("__exception_type")
                .map(|_| o.properties.get("message").cloned()),
        )
    };
    if let Some(message) = exc_message {
        let msg = match message {
            Some(Value::String(s)) => s.to_string(),
            Some(Value::Null) | Some(Value::Undefined) | None => String::new(),
            Some(other) => format!("{}", other),
        };
        return Value::String(Arc::from(msg.as_str()));
    }
    // Class instances with a `__type` tag get the spec-shaped
    // `[object <Name>]` rather than `[object]` (the Vybe Display
    // default for Ordinary).
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
    let Some(Value::Object(obj)) = arg else {
        return None;
    };
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

fn dispatch_weakmap(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    match method {
        "get" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(key, Value::Object(_)) {
                return Value::Undefined;
            }
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
            if !matches!(key, Value::Object(_)) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Invalid value used as weak map key",
                ));
                return Value::Null;
            }
            // find or insert
            let existing = {
                let m = obj.lock().unwrap();
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref keys) = ko.kind {
                        wm_key_ptr_find(keys, &key)
                    } else {
                        None
                    }
                } else {
                    None
                }
            };
            let mut m = obj.lock().unwrap();
            if let Some(pos) = existing {
                if let ObjectKind::Array(ref mut values) = m.kind {
                    values[pos] = val;
                }
            } else {
                if let ObjectKind::Array(ref mut values) = m.kind {
                    values.push(val);
                }
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP).cloned() {
                    let mut ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref mut keys) = ko.kind {
                        keys.push(key);
                    }
                }
            }
            drop(m);
            Value::Object(obj)
        }
        "has" => {
            let key = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(key, Value::Object(_)) {
                return Value::Bool(false);
            }
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
            if !matches!(key, Value::Object(_)) {
                return Value::Bool(false);
            }
            let mut m = obj.lock().unwrap();
            let pos = if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                let ko = keys_obj.lock().unwrap();
                if let ObjectKind::Array(ref keys) = ko.kind {
                    wm_key_ptr_find(keys, &key)
                } else {
                    None
                }
            } else {
                None
            };
            if let Some(pos) = pos {
                if let ObjectKind::Array(ref mut values) = m.kind {
                    values.remove(pos);
                }
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP).cloned() {
                    let mut ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref mut keys) = ko.kind {
                        keys.remove(pos);
                    }
                }
                return Value::Bool(true);
            }
            Value::Bool(false)
        }
        _ => Value::Undefined,
    }
}

// ── WeakSet dynamic dispatch ──────────────────────────────────────────────────

fn dispatch_weakset(
    ctx: &mut HostContext,
    obj: Arc<Mutex<Object>>,
    method: &str,
    args: &[Value],
) -> Value {
    match method {
        "add" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(v, Value::Object(_)) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Invalid value used in weak set",
                ));
                return Value::Null;
            }
            let mut so = obj.lock().unwrap();
            let already = if let ObjectKind::Array(ref vs) = so.kind {
                wm_key_ptr_find(vs, &v).is_some()
            } else {
                false
            };
            if !already {
                if let ObjectKind::Array(ref mut vs) = so.kind {
                    vs.push(v);
                }
            }
            drop(so);
            Value::Object(obj)
        }
        "has" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(v, Value::Object(_)) {
                return Value::Bool(false);
            }
            let so = obj.lock().unwrap();
            if let ObjectKind::Array(ref vs) = so.kind {
                return Value::Bool(wm_key_ptr_find(vs, &v).is_some());
            }
            Value::Bool(false)
        }
        "delete" => {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            if !matches!(v, Value::Object(_)) {
                return Value::Bool(false);
            }
            let mut so = obj.lock().unwrap();
            let pos = if let ObjectKind::Array(ref vs) = so.kind {
                wm_key_ptr_find(vs, &v)
            } else {
                None
            };
            if let Some(pos) = pos {
                if let ObjectKind::Array(ref mut vs) = so.kind {
                    vs.remove(pos);
                }
                return Value::Bool(true);
            }
            Value::Bool(false)
        }
        _ => Value::Undefined,
    }
}
