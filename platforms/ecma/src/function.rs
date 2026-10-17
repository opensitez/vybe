//! ECMA-262 §20.2 — Function.prototype.{bind, call, apply}.
//!
//! `bind` returns a new function ref carrying `__bound_args` so the VM
//! call dispatch in `vybe_runtime/src/calls.rs` prepends them on every
//! invocation. The VM hook is the same mechanism `new Promise(executor)`
//! uses for resolve/reject thunks — see `crate::promise`.
//!
//! `call(thisArg, ...args)` and `apply(thisArg, argsArray)` are the
//! spread/forward primitives. The Vybe call dispatch already passes the
//! receiver as the first arg of the args slice, so `call` and `apply`
//! just route through `ctx.invoke` with the resolved args list.

use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

static FUNCTION_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
const APPLY_INLINE_ARG_LIMIT: usize = 8;

#[inline]
fn invoke_bound_host_key() -> &'static (String, String) {
    static KEY: OnceLock<(String, String)> = OnceLock::new();
    KEY.get_or_init(|| ("ecma:function".to_string(), "invokeBound".to_string()))
}

#[inline]
fn to_string_host_key() -> &'static (String, String) {
    static KEY: OnceLock<(String, String)> = OnceLock::new();
    KEY.get_or_init(|| ("ecma:function".to_string(), "toString".to_string()))
}

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

pub fn shared_function_prototype() -> Value {
    Value::Object(
        FUNCTION_PROTOTYPE
            .get_or_init(|| {
                // §20.2.3: %Function.prototype% has own non-enumerable
                // `length: 0` and `name: ""` data properties — the
                // intrinsic kind prototypes (%AsyncFunction.prototype% …)
                // inherit them through their [[Prototype]] link.
                let mut proto = Object::new();
                proto.properties.reserve(3);
                proto.properties.insert("length".into(), Value::F64(0.0));
                proto
                    .properties
                    .insert("name".into(), crate::keys::string_value(""));
                proto.properties.insert(
                    "__nonenum".into(),
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                        crate::keys::string_value("length"),
                        crate::keys::string_value("name"),
                    ]))),
                );
                vybe_runtime::heap::alloc(proto)
            })
            .clone(),
    )
}

pub fn register(vm: &mut VM) {
    crate::perf::register_free_fn(
        vm,
        "ecma:function",
        "invokeBound",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Captures start AFTER the receiver slot. The VM places a host
            // callee's receiver FIRST, ahead of the bound captures, precisely
            // because this wrapper's capture count varies
            // (`[target, this, proto, ...partials]`) and a receiver appended
            // after them could not be located at all.
            // ⛔ THE CAPTURES START AT 0. §20.2.3.2: a bound function
            // IGNORES the thisArg of a later call — it already closed over
            // one — so its type declares no receiver
            // (`register_free_fn`), no slot is filled for it, and there is
            // no offset to compute. This used to read `receiver_argc()`,
            // asking the CALL a question only the callee's type can answer.
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            let bound_this = args.get(1).cloned().unwrap_or(Value::Undefined);
            let target_proto = args.get(2).cloned().unwrap_or(Value::Undefined);

            // The trailing arguments are a REST parameter — empty when the
            // caller passes fewer. `&args[3..]` panicked instead, and a host
            // panic takes down the whole worker, not just the call.
            //
            // ⛔ `user_args`, NOT `args.get(3..)`. The three leading values are
            // this wrapper's own captures; under
            // `ReceiverBinding::UniversalParameter` the CALL then puts a
            // receiver after them, and slicing from a fixed 3 handed that
            // receiver to the target as its first real argument. Measured:
            // `f.bind(o, 1)(2)` returned NaN and a bound method stored on an
            // object returned `11[object Object]`. `user_args` skips the
            // captures and the receiver slot, and is a no-op under the ambient
            // binding.
            // Everything past the three captures is `partials ++ callArgs`,
            // which is exactly what §20.2.3.2 hands the target. The receiver
            // sits BEFORE the captures, so it is already skipped by `base`.
            invoke_bound_target(
                ctx,
                target,
                bound_this,
                target_proto,
                args.get(3..).unwrap_or(&[]),
            )
        }),
    );

    let invoke_bound_idx = *vm
        .host_registry
        .get(invoke_bound_host_key())
        .expect("ecma:function.invokeBound must be registered before bind");

    // Function.prototype.name — §20.2.3.3: returns the name property.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "name",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let name = o
                    .properties
                    .get("name")
                    .or_else(|| o.properties.get("__fn_name"));
                if let Some(Value::String(s)) = name {
                    return Value::String(s.clone());
                }
            }
            crate::keys::string_value("")
        }),
    );

    // Function.prototype.length — §20.2.3.2: formal parameter count.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "length",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let len = o
                    .properties
                    .get("length")
                    .or_else(|| o.properties.get("__fn_arity"));
                match len {
                    Some(Value::I32(n)) => return Value::I32(*n),
                    Some(Value::F64(n)) => return Value::I32(*n as i32),
                    Some(Value::I64(n)) => return Value::I32(*n as i32),
                    _ => {}
                }
                if let ObjectKind::Function(f) = &o.kind {
                    return Value::I32(f.arity as i32);
                }
            }
            Value::I32(0)
        }),
    );

    // Function.prototype.toString — §20.2.3.5.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let name = o
                    .properties
                    .get("name")
                    .or_else(|| o.properties.get("__fn_name"))
                    .and_then(|v| {
                        if let Value::String(s) = v {
                            Some(s.to_string())
                        } else {
                            None
                        }
                    })
                    .unwrap_or_default();
                return Value::String(crate::keys::concat3_arc(
                    "function ",
                    &name,
                    "() { [native code] }",
                ));
            }
            crate::keys::string_value("function () { [native code] }")
        }),
    );

    // new Function(body) — §20.2.1.1: creates a callable from a body string.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let body = args
                .first()
                .map(crate::keys::value_display_string)
                .unwrap_or_default();
            let mut obj = Object::new();
            obj.properties.reserve(4);
            obj.properties
                .insert("name".into(), crate::keys::string_value("anonymous"));
            obj.properties.insert("length".into(), Value::I32(0));
            obj.properties
                .insert("__fn_body".into(), owned_string_value(body));
            obj.properties
                .insert("__fn_return".into(), Value::Undefined);
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // new Function(params, body) — §20.2.1.1 with parameters.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "newWithParams",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let params = args
                .first()
                .map(crate::keys::value_display_string)
                .unwrap_or_default();
            let body = args
                .get(1)
                .map(crate::keys::value_display_string)
                .unwrap_or_default();
            let mut obj = Object::new();
            obj.properties.reserve(5);
            obj.properties
                .insert("name".into(), crate::keys::string_value("anonymous"));
            let has_params = !params.is_empty();
            obj.properties
                .insert("length".into(), Value::I32(if has_params { 1 } else { 0 }));
            obj.properties
                .insert("__fn_params".into(), owned_string_value(params));
            obj.properties
                .insert("__fn_body".into(), owned_string_value(body));
            obj.properties
                .insert("__fn_return".into(), Value::Undefined);
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // bindWithArgs(fn, thisArg, ...args) — like bind with pre-supplied args.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "bindWithArgs",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            bind_function_with_arity(target, args.get(1..).unwrap_or(&[]), invoke_bound_idx)
        }),
    );

    // Function.prototype.bind(this_fn, thisArg, ...boundArgs) → new Function
    //
    // The returned function ref carries `__bound_args = [thisArg, ...boundArgs]`
    // and points at the same host fn idx as the receiver (or the same
    // chunk for user functions).
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "bind",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            bind_function(target, args.get(1..).unwrap_or(&[]), invoke_bound_idx)
        }),
    );

    // Initialize an existing derived receiver using a constructor supplied by
    // a separately compiled module. Constructor helpers use a trailing receiver
    // after their declared parameters, unlike ordinary method invocation.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "initialize",
        Box::new(|ctx, args| {
            let receiver = args.get(1).cloned().unwrap_or(Value::Undefined);
            let Some(Value::Object(class)) = args.first() else {
                return receiver;
            };
            let metadata = {
                let class = class.lock().unwrap();
                class
                    .properties
                    .get("__vybe_constructor_initializer")
                    .cloned()
                    .zip(
                        class
                            .properties
                            .get("__vybe_constructor_arity")
                            .map(|value| value.as_f64() as usize),
                    )
            };
            let Some((initializer, arity)) = metadata else {
                return receiver;
            };
            let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] =
                std::array::from_fn(|_| Value::Undefined);
            if let Some(length) = collect_initializer_inline(args.get(2), arity, &mut inline) {
                inline[arity] = receiver;
                return ctx.invoke(&initializer, &inline[..length]);
            }
            let mut packed = args.get(2).map(collect_apply_args).unwrap_or_default();
            packed.resize(arity, Value::Undefined);
            packed.push(receiver);
            ctx.invoke(&initializer, &packed)
        }),
    );

    // Function.prototype.call(this_fn, thisArg, ...args) → result
    //
    // Synchronously invokes the receiver with the given thisArg + args.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "call",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            // §20.2.3.3 `Function.prototype.call ( thisArg, ...args )` — `args`
            // is a REST parameter, so omitting it means EMPTY, not a panic.
            // `Function.prototype.call.call("x")` passes one argument and
            // `&args[2..]` panicked the worker; the two arguments above were
            // already guarded and this one was not.
            invoke_with_explicit_this(ctx, target, this_arg, args.get(2..).unwrap_or(&[]))
        }),
    );

    // Function.prototype.apply(this_fn, thisArg, argsArray) → result
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            let (callable, host_callee) = apply_target_kind(target);
            if !callable {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Function.apply target is not callable",
                ));
                return Value::Undefined;
            }
            let mut inline: [Value; APPLY_INLINE_ARG_LIMIT] =
                std::array::from_fn(|_| Value::Undefined);
            let apply_arg = args.get(2).unwrap_or(&undefined);
            let collected = if matches!(apply_arg, Value::Null | Value::Undefined) {
                Ok(CollectedApplyArgs::Inline(&inline[..0]))
            } else {
                collect_apply_args_once(ctx, Some(apply_arg), &mut inline)
            };
            let invoke_args = match collected {
                Ok(values) => values,
                Err(error) => {
                    ctx.throw_value(error);
                    return Value::Undefined;
                }
            };
            // ⛔ A HOST BUILTIN HAS NO `this` — IT READS ARGUMENT 0.
            //
            // `apply` hands `thisArg` to the callee's `this`, which is right
            // for a bytecode function and useless for `ecma:array.slice`, whose
            // body opens `array_of(args, 0)`. Under the ambient binding that
            // still worked, because `getMethodForCall(arr, "slice", bind=true)`
            // had stuffed the receiver into `__bound_args` and the VM prepended
            // it. Under `ReceiverAbi::Parameter` I pass `bind = false` — binding
            // is that untypeable hidden-prepend channel — so nothing prepends it
            // and `slice` read the START index as its array: `[1,2,3,4].slice(1)`
            // returned `[]`, `indexOf` returned -1.
            //
            // `invoke_with_receiver` is exactly this split: it PREPENDS for a
            // host callee (which reads argument 0) and SETS the channel for a
            // bytecode one (which gets it prepended by `invoke`), so neither
            // ends up with two. Under `Ambient` it binds the global and behaves
            // as before.
            // ⛔ ONLY A HOST CALLEE GOES THROUGH `invoke_with_receiver`.
            //
            // A host builtin reads its receiver as ARGUMENT 0, so it needs the
            // prepend. A BYTECODE callee needs `invoke_compiled_function`,
            // which also does REST PACKING and builds `arguments` — routing it
            // through the raw invoke skipped both: `len.apply(null, [1,2])`
            // reported `arguments.length` as `undefined`, and an EMPTY args
            // array threw "Cannot convert undefined or null to object".
            invoke_apply_collected_args(ctx, target, this_arg, host_callee, invoke_args.as_slice())
        }),
    );

    // §10.2.9 SetFunctionName for runtime-computed property keys:
    // anonymous functions assigned under a computed key take the key's
    // string form; symbol keys become "[<description>]" (or "" when the
    // symbol has none). Already-named functions keep their name.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "setFunctionName",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(f)) = args.first() {
                let undefined = Value::Undefined;
                let key = args.get(1).unwrap_or(&undefined);
                let mut o = f.lock().unwrap();
                let anonymous = match o.properties.get("name") {
                    Some(Value::String(n)) => n.is_empty() || n.starts_with("__anon_fn_"),
                    _ => true,
                };
                if anonymous {
                    let name = match key {
                        Value::Symbol(s) => {
                            if crate::symbol::has_description(s) {
                                format!("[{}]", s)
                            } else {
                                String::new()
                            }
                        }
                        other => crate::keys::value_display_string(other),
                    };
                    o.properties.insert("name".into(), owned_string_value(name));
                }
            }
            Value::Undefined
        }),
    );

    // §20.2.3.5 Function.prototype.toString. Source text isn't retained,
    // so every form uses the spec's NativeFunction fallback shape with the
    // function's kind classifier tokens (async / * / =>) and name.
    crate::perf::register_host_fn(
        vm,
        "ecma:function",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            owned_string_value(function_to_string(target))
        }),
    );
    let to_string_idx = *vm
        .host_registry
        .get(to_string_host_key())
        .expect("ecma:function.toString just registered");
    if let Value::Object(proto) = shared_function_prototype() {
        let mut p = proto.lock().unwrap();
        if !p.properties.contains_key("toString") {
            let mut ts = Object::new();
            ts.kind = ObjectKind::HostFunction(to_string_idx);
            ts.properties
                .insert("name".into(), crate::keys::string_value("toString"));
            ts.properties.insert("length".into(), Value::F64(0.0));
            ts.properties
                .insert("__vybe_method_receiver".into(), Value::Bool(true));
            p.properties.insert(
                "toString".into(),
                Value::Object(vybe_runtime::heap::alloc(ts)),
            );
            // toString is non-enumerable on %Function.prototype%.
            if let Some(Value::Object(ne)) = p.properties.get("__nonenum") {
                let mut a = ne.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = a.kind {
                    elems.push(crate::keys::string_value("toString"));
                }
            }
        }
    }
}

// Initialize the caller's buffer with Undefined before calling. The final
// slot is reserved for the receiver, replacing any excess collected argument.
fn collect_initializer_inline(
    value: Option<&Value>,
    arity: usize,
    inline: &mut [Value; APPLY_INLINE_ARG_LIMIT + 1],
) -> Option<usize> {
    if arity > APPLY_INLINE_ARG_LIMIT {
        return None;
    }
    collect_apply_args_inline(value, inline)?;
    Some(arity + 1)
}

/// §20.2.3.5 — synthesized function string. Bound and native functions use
/// the NativeFunction form; compiled functions add the kind tokens the
/// compiler stamped (`__fn_kind`, `__fn_arrow`).
fn function_to_string(target: &Value) -> String {
    let Value::Object(obj) = target else {
        return "function () { [native code] }".to_string();
    };
    let o = obj.lock().unwrap();
    // §20.2.3.5 step 2 note: bound function exotic objects stringify as
    // native — they never expose their target's source.
    if o.properties.contains_key("__bound_args") {
        return "function () { [native code] }".to_string();
    }
    let name = match o.properties.get("name") {
        Some(Value::String(n)) => n.to_string(),
        _ => String::new(),
    };
    if matches!(o.kind, ObjectKind::HostFunction(_)) {
        return format!("function {}() {{ [native code] }}", name);
    }
    if matches!(o.properties.get("__fn_arrow"), Some(Value::Bool(true))) {
        let is_async = matches!(
            o.properties.get("__fn_kind"),
            Some(Value::String(k)) if k.as_ref() == "async"
        );
        return format!(
            "{}() => {{ [native code] }}",
            if is_async { "async " } else { "" }
        );
    }
    match o.properties.get("__fn_kind") {
        Some(Value::String(k)) if k.as_ref() == "async" => {
            format!("async function {}() {{ [native code] }}", name)
        }
        Some(Value::String(k)) if k.as_ref() == "generator" => {
            format!("function* {}() {{ [native code] }}", name)
        }
        Some(Value::String(k)) if k.as_ref() == "async_generator" => {
            format!("async function* {}() {{ [native code] }}", name)
        }
        _ => format!("function {}() {{ [native code] }}", name),
    }
}

pub fn invoke_bound_callback_if_needed(
    ctx: &mut HostContext,
    callback: &Value,
    args: &[Value],
) -> Option<Value> {
    let prepared = prepare_bound_callback(callback)?;
    Some(invoke_prepared_bound_callback(ctx, &prepared, args))
}

#[derive(Clone)]
pub struct PreparedBoundCallback {
    target: Value,
    bound_this: Value,
    target_proto: Value,
    partials: Vec<Value>,
}

#[cfg(test)]
#[derive(Clone)]
enum BoundPartials {
    Inline { values: [Value; APPLY_INLINE_ARG_LIMIT], len: usize },
    Heap(Vec<Value>),
}

#[cfg(test)]
impl BoundPartials {
    fn snapshot(values: &[Value]) -> Self {
        if values.len() <= APPLY_INLINE_ARG_LIMIT {
            let mut inline = std::array::from_fn(|_| Value::Undefined);
            inline[..values.len()].clone_from_slice(values);
            Self::Inline { values: inline, len: values.len() }
        } else {
            Self::Heap(values.to_vec())
        }
    }
}

#[cfg(test)]
impl std::ops::Deref for BoundPartials {
    type Target = [Value];

    fn deref(&self) -> &Self::Target {
        match self {
            Self::Inline { values, len } => &values[..*len],
            Self::Heap(values) => values,
        }
    }
}

#[cfg(test)]
mod bound_capture_storage_tests {
    use super::*;

    #[test]
    fn captures_preserve_order_identity_and_snapshot_at_storage_boundary() {
        for len in [0, 1, 3, 8, 9, 32] {
            let identity = vybe_runtime::heap::alloc(Object::new());
            let mut values: Vec<Value> = (0..len).map(|n| Value::I32(n as i32)).collect();
            if len != 0 { values[0] = Value::Object(identity.clone()); }
            let snapshot = BoundPartials::snapshot(&values);
            let copy = snapshot.clone();
            assert_eq!(snapshot.len(), len);
            assert_eq!(matches!(snapshot, BoundPartials::Inline { .. }), len <= APPLY_INLINE_ARG_LIMIT);
            values.clear();
            for captured in [&*snapshot, &*copy] {
                assert_eq!(captured.len(), len);
                for (index, value) in captured.iter().enumerate() {
                    if index == 0 {
                        let Value::Object(actual) = value else { panic!("identity missing") };
                        assert!(Arc::ptr_eq(actual, &identity));
                    } else {
                        assert!(matches!(value, Value::I32(n) if *n == index as i32));
                    }
                }
            }
        }
    }

    #[test]
    fn decoding_releases_locks_and_does_not_retain_mutable_capture_storage() {
        for len in [0, 1, 8, 9, 32] {
            let mut captures = vec![Value::I32(101), Value::I32(102), Value::I32(103)];
            captures.extend((0..len).map(|n| Value::I32(n as i32)));
            let backing = vybe_runtime::heap::alloc(Object::new_array(captures));
            let mut wrapper = Object::new();
            wrapper.properties.insert("name".into(), crate::keys::string_value("bound capture"));
            wrapper.properties.insert("__bound_args".into(), Value::Object(backing.clone()));
            let wrapper = vybe_runtime::heap::alloc(wrapper);
            let decoded = decode_bound_callback(&Value::Object(wrapper.clone())).unwrap();
            assert!(wrapper.try_lock().is_ok());
            backing.lock().unwrap().kind = ObjectKind::Array(Vec::new());
            assert!(matches!(decoded.target, Value::I32(101)));
            assert!(matches!(decoded.bound_this, Value::I32(102)));
            assert!(matches!(decoded.target_proto, Value::I32(103)));
            assert_eq!(decoded.partials.len(), len);
            for (index, value) in decoded.partials.iter().enumerate() {
                assert!(matches!(value, Value::I32(n) if *n == index as i32));
            }
        }
    }

    #[test]
    #[ignore = "paired bound capture snapshot kernel, not complete bind or guest timing"]
    fn paired_capture_snapshot_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        for len in [0, 1, 3, 8, 9, 32] {
            let values: Vec<Value> = (0..len).map(|_| crate::keys::string_value("captured")).collect();
            let mut previous = Duration::ZERO;
            let mut current = Duration::ZERO;
            for round in 0..8 {
                for optimized in if round % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..20_000 {
                        if optimized {
                            black_box(BoundPartials::snapshot(black_box(&values)));
                        } else {
                            black_box(black_box(&values).iter().cloned().collect::<Vec<Value>>());
                        }
                    }
                    if optimized { current += start.elapsed(); } else { previous += start.elapsed(); }
                }
            }
            eprintln!("bound captures len={len}: previous={previous:?}, current={current:?}");
        }
    }
}

fn decode_bound_callback(callback: &Value) -> Option<PreparedBoundCallback> {
    let Value::Object(obj) = callback else {
        return None;
    };

    let object = obj.lock().unwrap();
    match object.properties.get("__bound_args") {
        Some(Value::Object(bound)) => {
            if !matches!(object.properties.get("name"), Some(Value::String(text)) if text.starts_with("bound "))
            {
                return None;
            }
            let bound_object = bound.lock().unwrap();
            if let ObjectKind::Array(values) = &bound_object.kind {
                if values.len() < 3 {
                    return None;
                }
                Some(PreparedBoundCallback {
                    target: values[0].clone(),
                    bound_this: values[1].clone(),
                    target_proto: values[2].clone(),
                    partials: values.iter().skip(3).cloned().collect(),
                })
            } else {
                None
            }
        }
        _ => None,
    }
}

pub fn prepare_bound_callback(callback: &Value) -> Option<PreparedBoundCallback> {
    decode_bound_callback(callback)
}

pub fn invoke_prepared_bound_callback(
    ctx: &mut HostContext,
    prepared: &PreparedBoundCallback,
    args: &[Value],
) -> Value {
    if prepared.partials.is_empty() {
        return invoke_bound_target(
            ctx,
            &prepared.target,
            prepared.bound_this.clone(),
            prepared.target_proto.clone(),
            args,
        );
    }
    let total_len = prepared.partials.len() + args.len();
    if total_len <= APPLY_INLINE_ARG_LIMIT {
        let mut inline: [Value; APPLY_INLINE_ARG_LIMIT] = std::array::from_fn(|_| Value::Undefined);
        let mut index = 0;
        for value in prepared.partials.iter() {
            inline[index] = value.clone();
            index += 1;
        }
        for value in args {
            inline[index] = value.clone();
            index += 1;
        }
        return invoke_bound_target(
            ctx,
            &prepared.target,
            prepared.bound_this.clone(),
            prepared.target_proto.clone(),
            &inline[..total_len],
        );
    }
    let mut invoke_args = Vec::with_capacity(prepared.partials.len() + args.len());
    invoke_args.extend(prepared.partials.iter().cloned());
    invoke_args.extend_from_slice(args);
    invoke_bound_target(
        ctx,
        &prepared.target,
        prepared.bound_this.clone(),
        prepared.target_proto.clone(),
        &invoke_args,
    )
}

pub fn try_invoke_bound_callback_if_needed(
    ctx: &mut HostContext,
    callback: &Value,
    args: &[Value],
) -> Option<Result<Value, Value>> {
    let prepared = decode_bound_callback(callback)?;
    if prepared.partials.is_empty() {
        return Some(try_invoke_bound_target(
            ctx,
            &prepared.target,
            prepared.bound_this.clone(),
            prepared.target_proto.clone(),
            args,
        ));
    }
    let total_len = prepared.partials.len() + args.len();
    if total_len <= APPLY_INLINE_ARG_LIMIT {
        let mut inline: [Value; APPLY_INLINE_ARG_LIMIT] = std::array::from_fn(|_| Value::Undefined);
        let mut index = 0;
        for value in prepared.partials.iter() {
            inline[index] = value.clone();
            index += 1;
        }
        for value in args {
            inline[index] = value.clone();
            index += 1;
        }
        return Some(try_invoke_bound_target(
            ctx,
            &prepared.target,
            prepared.bound_this.clone(),
            prepared.target_proto.clone(),
            &inline[..total_len],
        ));
    }
    let mut invoke_args = Vec::with_capacity(prepared.partials.len() + args.len());
    invoke_args.extend(prepared.partials.iter().cloned());
    invoke_args.extend_from_slice(args);
    Some(try_invoke_bound_target(
        ctx,
        &prepared.target,
        prepared.bound_this.clone(),
        prepared.target_proto.clone(),
        &invoke_args,
    ))
}

pub fn invoke_with_explicit_this(
    ctx: &mut HostContext,
    target: &Value,
    this_arg: Value,
    args: &[Value],
) -> Value {
    // §10.5.12 [[Call]] on a proxy: the apply trap fires with thisArg
    // (or the target is invoked with it when trapless). Reaches here via
    // Function.prototype.call/apply/bind on proxy-wrapped functions.
    match explicit_this_call_kind(target) {
        ExplicitThisCallKind::Proxy(proxy_target, handler) => {
            if let Some(trap) = crate::object::proxy_trap(&handler, "apply") {
                let args_arr = make_arguments_array(args);
                return invoke_with_explicit_this(
                    ctx,
                    &trap,
                    handler,
                    &[proxy_target, this_arg, args_arr],
                );
            }
            return invoke_with_explicit_this(ctx, &proxy_target, this_arg, args);
        }
        ExplicitThisCallKind::BoundHost => {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = ctx.invoke(target, args);
            ctx.set_js_this(previous_this);
            result
        }
        ExplicitThisCallKind::Compiled(fixed_count) => {
            // Arrows need no special case here: they capture lexical
            // `this` at creation (compiler-emitted upvalue) and never read
            // the ambient binding this sets (§10.2.11).
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = invoke_compiled_function_with_arity(ctx, target, args, fixed_count);
            ctx.set_js_this(previous_this);
            result
        }
        ExplicitThisCallKind::Constant(value) => value,
        call_kind => {
            // ⛔ DO NOT PREPEND, AND DO NOT BRANCH ON THE CALLEE'S KIND.
            // The receiver slot of a host callee is filled in exactly ONE
            // place — `call_value_inner` — so prepending here handed the
            // callee two and shifted every real argument. Branching on whether
            // THIS host function "uses an explicit receiver" is the deeper
            // problem: it makes a call's meaning depend on the callee's
            // runtime kind, which no WASM call instruction can express, and a
            // funcref's type cannot say "argument 0 is a receiver, sometimes".
            //
            // Binding the channel is the whole job. `invoke_with_receiver`
            // says WHICH receiver; the VM decides WHERE it goes, once, for
            // every callee alike.
            // Under `ReceiverAbi::Parameter` every host callee's argument 0
            // is its receiver, so binding is unconditional and
            // `invoke_with_receiver` adds nothing to the list — the VM fills
            // the slot.
            //
            // ⛔ UNDER THE AMBIENT ABI IT IS NOT UNCONDITIONAL. There the
            // receiver is only placed for a host function that actually
            // DECLARES one; handing it to every host callee shifts the
            // arguments of the ones that do not — measured on csharp as
            // `HashSet.SetEquals` answering False. The ambient surface was
            // never uniform, and pretending it is breaks it.
            if ctx.receiver_is_parameter() {
                // Under `Parameter` the VM fills the receiver slot; binding the
                // channel is all this has to do.
                ctx.invoke_with_receiver(target, this_arg, args)
            } else if matches!(call_kind, ExplicitThisCallKind::Host(true)) {
                // ⛔ PLACE THE RECEIVER, DO NOT REBIND THE CHANNEL.
                // `invoke_with_receiver` brackets the call with its own
                // receiver binding, and rebinding disturbs an ENCLOSING
                // method's own receiver. Prepending and leaving the channel
                // alone is what component-class method bodies depend on.
                invoke_with_prepended_receiver(ctx, target, this_arg, args)
            } else {
                ctx.invoke(target, args)
            }
        }
    }
}

pub fn try_invoke_with_explicit_this(
    ctx: &mut HostContext,
    target: &Value,
    this_arg: Value,
    args: &[Value],
) -> Result<Value, Value> {
    match explicit_this_call_kind(target) {
        ExplicitThisCallKind::Proxy(proxy_target, handler) => {
            if let Some(trap) = crate::object::proxy_trap(&handler, "apply") {
                let args_arr = make_arguments_array(args);
                return try_invoke_with_explicit_this(
                    ctx,
                    &trap,
                    handler,
                    &[proxy_target, this_arg, args_arr],
                );
            }
            return try_invoke_with_explicit_this(ctx, &proxy_target, this_arg, args);
        }
        ExplicitThisCallKind::BoundHost => {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = ctx.try_invoke(target, args);
            ctx.set_js_this(previous_this);
            result
        }
        ExplicitThisCallKind::Compiled(fixed_count) => {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = try_invoke_compiled_function_with_arity(ctx, target, args, fixed_count);
            ctx.set_js_this(previous_this);
            result
        }
        ExplicitThisCallKind::Constant(value) => Ok(value),
        call_kind => {
            if ctx.receiver_is_parameter() {
                // Match the non-fallible path: the VM places the receiver
                // under the Parameter ABI. Prepending here would duplicate
                // it and shift the user's arguments.
                ctx.try_invoke_with_receiver(target, this_arg, args)
            } else if matches!(call_kind, ExplicitThisCallKind::Host(true)) {
                try_invoke_with_prepended_receiver(ctx, target, this_arg, args)
            } else {
                ctx.try_invoke(target, args)
            }
        }
    }
}

fn invoke_bound_target(
    ctx: &mut HostContext,
    target: &Value,
    bound_this: Value,
    target_proto: Value,
    args: &[Value],
) -> Value {
    match target {
        Value::Object(target_obj)
            if matches!(target_obj.lock().unwrap().kind, ObjectKind::Function(_)) =>
        {
            let previous_this = ctx.current_js_this();
            let constructor_call = matches!((&previous_this, &target_proto),
                (Value::Object(current), Value::Object(expected_proto))
                    if matches!(current.lock().unwrap().properties.get("__proto__"), Some(Value::Object(proto)) if Arc::ptr_eq(proto, expected_proto))
            );
            if !constructor_call {
                ctx.set_js_this(bound_this);
            }
            let result = invoke_compiled_function(ctx, target, args);
            ctx.set_js_this(previous_this.clone());
            if constructor_call && !matches!(result, Value::Object(_)) {
                previous_this
            } else {
                result
            }
        }
        _ => {
            // ⛔ DO NOT BRANCH ON THE CALLEE'S KIND. Under
            // `ReceiverAbi::Parameter` argument 0 of EVERY host callee is its
            // receiver, so `invoke_with_receiver` says which one and the VM
            // places it — one signature, one meaning. Asking whether this
            // particular host function "uses an explicit receiver" is a
            // runtime type test at the call, which no WASM call instruction
            // can express: measured, `Array.prototype.join.bind([1,2,3])("-")`
            // answered EMPTY because `join` carries no receiver marker, so the
            // channel was never bound and the VM filled the slot from a stale
            // receiver. §20.2.3.2 hands the target the BOUND this, always.
            //
            // The ambient arm is unchanged: there the VM places nothing, so
            // only a host function that declares a receiver may be given one.
            if ctx.receiver_is_parameter() {
                ctx.invoke_with_receiver(target, bound_this, args)
            } else if host_function_uses_explicit_receiver(target) {
                invoke_with_prepended_receiver(ctx, target, bound_this, args)
            } else {
                ctx.invoke(target, args)
            }
        }
    }
}

fn try_invoke_bound_target(
    ctx: &mut HostContext,
    target: &Value,
    bound_this: Value,
    target_proto: Value,
    args: &[Value],
) -> Result<Value, Value> {
    match target {
        Value::Object(target_obj)
            if matches!(target_obj.lock().unwrap().kind, ObjectKind::Function(_)) =>
        {
            let previous_this = ctx.current_js_this();
            let constructor_call = matches!((&previous_this, &target_proto),
                (Value::Object(current), Value::Object(expected_proto))
                    if matches!(current.lock().unwrap().properties.get("__proto__"), Some(Value::Object(proto)) if Arc::ptr_eq(proto, expected_proto))
            );
            if !constructor_call {
                ctx.set_js_this(bound_this);
            }
            let result = try_invoke_compiled_function(ctx, target, args);
            ctx.set_js_this(previous_this.clone());
            match result {
                Ok(value) if constructor_call && !matches!(value, Value::Object(_)) => {
                    Ok(previous_this)
                }
                other => other,
            }
        }
        _ => {
            // ⛔ DO NOT BRANCH ON THE CALLEE'S KIND. Under
            // `ReceiverAbi::Parameter` argument 0 of EVERY host callee is its
            // receiver, so `invoke_with_receiver` says which one and the VM
            // places it — one signature, one meaning. Asking whether this
            // particular host function "uses an explicit receiver" is a
            // runtime type test at the call, which no WASM call instruction
            // can express: measured, `Array.prototype.join.bind([1,2,3])("-")`
            // answered EMPTY because `join` carries no receiver marker, so the
            // channel was never bound and the VM filled the slot from a stale
            // receiver. §20.2.3.2 hands the target the BOUND this, always.
            //
            // The ambient arm is unchanged: there the VM places nothing, so
            // only a host function that declares a receiver may be given one.
            if ctx.receiver_is_parameter() {
                ctx.try_invoke_with_receiver(target, bound_this, args)
            } else if host_function_uses_explicit_receiver(target) {
                try_invoke_with_prepended_receiver(ctx, target, bound_this, args)
            } else {
                ctx.try_invoke(target, args)
            }
        }
    }
}

pub(crate) fn invoke_with_prepended_receiver(
    ctx: &mut HostContext,
    target: &Value,
    receiver: Value,
    args: &[Value],
) -> Value {
    if args.len() <= APPLY_INLINE_ARG_LIMIT {
        let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] =
            std::array::from_fn(|_| Value::Undefined);
        inline[0] = receiver;
        for (index, arg) in args.iter().enumerate() {
            inline[index + 1] = arg.clone();
        }
        ctx.invoke(target, &inline[..args.len() + 1])
    } else {
        let mut invoke_args = Vec::with_capacity(args.len() + 1);
        invoke_args.push(receiver);
        invoke_args.extend_from_slice(args);
        ctx.invoke(target, &invoke_args)
    }
}

fn try_invoke_with_prepended_receiver(
    ctx: &mut HostContext,
    target: &Value,
    receiver: Value,
    args: &[Value],
) -> Result<Value, Value> {
    if args.len() <= APPLY_INLINE_ARG_LIMIT {
        let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] =
            std::array::from_fn(|_| Value::Undefined);
        inline[0] = receiver;
        for (index, arg) in args.iter().enumerate() {
            inline[index + 1] = arg.clone();
        }
        ctx.try_invoke(target, &inline[..args.len() + 1])
    } else {
        let mut invoke_args = Vec::with_capacity(args.len() + 1);
        invoke_args.push(receiver);
        invoke_args.extend_from_slice(args);
        ctx.try_invoke(target, &invoke_args)
    }
}

fn invoke_compiled_function(ctx: &mut HostContext, target: &Value, args: &[Value]) -> Value {
    invoke_compiled_function_with_arity(ctx, target, args, compiled_rest_fixed_arity(target))
}

fn invoke_compiled_function_with_arity(
    ctx: &mut HostContext,
    target: &Value,
    args: &[Value],
    fixed_count: Option<usize>,
) -> Value {
    let Some(fixed_count) = fixed_count else {
        return ctx.invoke(target, args);
    };

    let total_len = fixed_count + 1;
    if total_len <= APPLY_INLINE_ARG_LIMIT {
        let mut packed_args: [Value; APPLY_INLINE_ARG_LIMIT] =
            std::array::from_fn(|_| Value::Undefined);
        for index in 0..fixed_count {
            packed_args[index] = args.get(index).cloned().unwrap_or(Value::Undefined);
        }
        packed_args[fixed_count] = make_rest_array(args, fixed_count);
        return ctx.invoke(target, &packed_args[..total_len]);
    }

    let mut packed_args = Vec::with_capacity(fixed_count + 1);
    for index in 0..fixed_count {
        packed_args.push(args.get(index).cloned().unwrap_or(Value::Undefined));
    }
    packed_args.push(make_rest_array(args, fixed_count));
    ctx.invoke(target, &packed_args)
}

fn try_invoke_compiled_function(
    ctx: &mut HostContext,
    target: &Value,
    args: &[Value],
) -> Result<Value, Value> {
    try_invoke_compiled_function_with_arity(ctx, target, args, compiled_rest_fixed_arity(target))
}

fn try_invoke_compiled_function_with_arity(
    ctx: &mut HostContext,
    target: &Value,
    args: &[Value],
    fixed_count: Option<usize>,
) -> Result<Value, Value> {
    let Some(fixed_count) = fixed_count else {
        return ctx.try_invoke(target, args);
    };

    let total_len = fixed_count + 1;
    if total_len <= APPLY_INLINE_ARG_LIMIT {
        let mut packed_args: [Value; APPLY_INLINE_ARG_LIMIT] =
            std::array::from_fn(|_| Value::Undefined);
        for index in 0..fixed_count {
            packed_args[index] = args.get(index).cloned().unwrap_or(Value::Undefined);
        }
        packed_args[fixed_count] = make_rest_array(args, fixed_count);
        return ctx.try_invoke(target, &packed_args[..total_len]);
    }

    let mut packed_args = Vec::with_capacity(fixed_count + 1);
    for index in 0..fixed_count {
        packed_args.push(args.get(index).cloned().unwrap_or(Value::Undefined));
    }
    packed_args.push(make_rest_array(args, fixed_count));
    ctx.try_invoke(target, &packed_args)
}

fn make_rest_array(args: &[Value], fixed_count: usize) -> Value {
    let rest_len = args.len().saturating_sub(fixed_count);
    let mut rest = Vec::with_capacity(rest_len);
    for value in args.iter().skip(fixed_count) {
        rest.push(value.clone());
    }
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(rest)))
}

fn compiled_rest_fixed_arity(target: &Value) -> Option<usize> {
    let Value::Object(obj) = target else {
        return None;
    };
    let object = obj.lock().unwrap();
    if !matches!(object.kind, ObjectKind::Function(_)) {
        return None;
    }
    rest_fixed_arity_from_object(&object)
}

fn rest_fixed_arity_from_object(object: &Object) -> Option<usize> {
    match object.properties.get("__vybe_rest_fixed_arity") {
        Some(Value::I32(value)) if *value >= 0 => Some(*value as usize),
        Some(Value::I64(value)) if *value >= 0 => Some(*value as usize),
        Some(Value::F64(value)) if *value >= 0.0 => Some(*value as usize),
        _ => None,
    }
}

fn host_function_uses_explicit_receiver(target: &Value) -> bool {
    let Value::Object(obj) = target else {
        return false;
    };
    let object = obj.lock().unwrap();
    matches!(object.kind, ObjectKind::HostFunction(_))
        && matches!(
            object.properties.get("__vybe_method_receiver"),
            Some(Value::Bool(true))
        )
}

enum ExplicitThisCallKind {
    Proxy(Value, Value),
    BoundHost,
    Compiled(Option<usize>),
    Constant(Value),
    Host(bool),
    Other,
}

// Snapshot all dispatch-only metadata under one lock. No snapshot survives
// this call, and the lock is released before proxy traps or user code run.
fn explicit_this_call_kind(target: &Value) -> ExplicitThisCallKind {
    let Value::Object(obj) = target else {
        return ExplicitThisCallKind::Other;
    };
    let object = obj.lock().unwrap();
    if let (Some(target), Some(handler)) = (
        object.properties.get("__vybe_proxy_target"),
        object.properties.get("__vybe_proxy_handler"),
    ) {
        return ExplicitThisCallKind::Proxy(target.clone(), handler.clone());
    }
    match &object.kind {
        ObjectKind::HostFunction(_) if object.properties.contains_key("__bound_args") => {
            ExplicitThisCallKind::BoundHost
        }
        ObjectKind::Function(_) => ExplicitThisCallKind::Compiled(rest_fixed_arity_from_object(&object)),
        _ if object.properties.contains_key("__fn_return") => {
            ExplicitThisCallKind::Constant(object.properties.get("__fn_return").unwrap().clone())
        }
        ObjectKind::HostFunction(_) => ExplicitThisCallKind::Host(matches!(
            object.properties.get("__vybe_method_receiver"),
            Some(Value::Bool(true))
        )),
        _ => ExplicitThisCallKind::Other,
    }
}

// Classify only the initial callability check. Argument getters may mutate
// callable metadata, so final dispatch must still inspect the target afresh.
fn apply_target_kind(target: &Value) -> (bool, bool) {
    let Value::Object(object) = target else {
        return (false, false);
    };
    let object = object.lock().unwrap();
    let callable = matches!(object.kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_));
    let plain_host = matches!(object.kind, ObjectKind::HostFunction(_))
        && !(object.properties.contains_key("__vybe_proxy_target")
            && object.properties.contains_key("__vybe_proxy_handler"));
    (callable, plain_host)
}

#[cfg(test)]
fn is_host_callee(target: &Value) -> bool {
    matches!(
        target,
        Value::Object(obj)
            if matches!(obj.lock().map(|g| {
                matches!(g.kind, ObjectKind::HostFunction(_))
                    && !(g.properties.contains_key("__vybe_proxy_target")
                        && g.properties.contains_key("__vybe_proxy_handler"))
            }), Ok(true))
    )
}

pub(crate) fn is_callable(value: &Value) -> bool {
    matches!(value, Value::Object(object)
        if matches!(object.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)))
}

fn invoke_apply_collected_args(
    ctx: &mut HostContext,
    target: &Value,
    this_arg: Value,
    host_callee: bool,
    invoke_args: &[Value],
) -> Value {
    if host_callee && ctx.receiver_is_parameter() {
        ctx.invoke_with_receiver(target, this_arg, invoke_args)
    } else {
        // Ambient host functions do not all declare a receiver. The shared
        // helper distinguishes receiver-bearing methods from free functions;
        // forwarding solely by HostFunction kind prepended an extra argument.
        invoke_with_explicit_this(ctx, target, this_arg, invoke_args)
    }
}

#[inline]
fn make_arguments_array(args: &[Value]) -> Value {
    if args.is_empty() {
        crate::array::make_array(Vec::new())
    } else {
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(args.to_vec())))
    }
}

fn apply_array_like_length(object: &Object) -> usize {
    apply_array_like_length_value(object.properties.get("length"))
}

fn apply_array_like_length_value(value: Option<&Value>) -> usize {
    match value {
        Some(Value::I32(value)) if *value > 0 => *value as usize,
        Some(Value::I64(value)) if *value > 0 => *value as usize,
        Some(Value::F64(value)) if *value > 0.0 => *value as usize,
        Some(Value::String(text)) => crate::keys::non_negative_integer_index_key(text).unwrap_or(0),
        _ => 0,
    }
}

pub(crate) enum CollectedApplyArgs<'a> {
    Inline(&'a [Value]),
    Heap(Vec<Value>),
}

impl CollectedApplyArgs<'_> {
    pub(crate) fn as_slice(&self) -> &[Value] {
        match self {
            Self::Inline(values) => values,
            Self::Heap(values) => values,
        }
    }
}

// Select inline storage or a heap snapshot under one source lock. Neither
// result borrows the source object, so callbacks may mutate it freely.
pub(crate) fn collect_apply_args_once<'a, const N: usize>(
    ctx: &mut HostContext,
    value: Option<&Value>,
    inline: &'a mut [Value; N],
) -> Result<CollectedApplyArgs<'a>, Value> {
    let Some(Value::Object(obj)) = value else {
        return Err(crate::error::new_error(
            ctx,
            "TypeError",
            "apply argument list must be an object",
        ));
    };
    let object = obj.lock().unwrap();
    let needs_get = object.properties.contains_key("__vybe_proxy_target")
        || object
            .properties
            .keys()
            .any(|key| key.starts_with("__get_"));
    if let ObjectKind::Array(values) = &object.kind {
        if !needs_get && !object.properties.contains_key("__holes") {
            if values.len() > N {
                return Ok(CollectedApplyArgs::Heap(values.clone()));
            }
            inline[..values.len()].clone_from_slice(values);
            return Ok(CollectedApplyArgs::Inline(&inline[..values.len()]));
        }
    }

    // Own-data lists have no callbacks while collecting. Missing properties,
    // accessors, sparse arrays and proxies must perform observable Get calls.
    if !needs_get
        && !matches!(object.kind, ObjectKind::Array(_))
        && object.properties.contains_key("length")
        && !matches!(object.properties.get("length"), Some(Value::Object(_)))
    {
        let length = apply_list_length(ctx, object.properties.get("length").unwrap())?;
        let element =
            |index| crate::keys::with_index_key(index, |key| object.properties.get(key).cloned());
        if length <= N {
            let mut complete = true;
            for (index, slot) in inline.iter_mut().enumerate().take(length) {
                match element(index) {
                    Some(value) => *slot = value,
                    None => {
                        complete = false;
                        break;
                    }
                }
            }
            if complete {
                return Ok(CollectedApplyArgs::Inline(&inline[..length]));
            }
        } else {
            let mut values = reserve_apply_list(ctx, length)?;
            for index in 0..length {
                match element(index) {
                    Some(value) => values.push(value),
                    None => break,
                }
            }
            if values.len() == length {
                return Ok(CollectedApplyArgs::Heap(values));
            }
        }
    }
    drop(object);
    let source = value.unwrap();
    let length_value = crate::reflect::try_reflect_get(ctx, source, "length", source.clone())?;
    let length = apply_list_length(ctx, &length_value)?;
    let mut heap_values = if length > N {
        reserve_apply_list(ctx, length)?
    } else {
        Vec::new()
    };
    let mut element = |index| {
        crate::keys::with_index_key(index, |key| {
            crate::reflect::try_reflect_get(ctx, source, key, source.clone())
        })
    };
    if length <= N {
        for (index, slot) in inline.iter_mut().enumerate().take(length) {
            *slot = element(index)?;
        }
        Ok(CollectedApplyArgs::Inline(&inline[..length]))
    } else {
        for index in 0..length {
            heap_values.push(element(index)?);
        }
        Ok(CollectedApplyArgs::Heap(heap_values))
    }
}

fn apply_list_length(ctx: &mut HostContext, value: &Value) -> Result<usize, Value> {
    let length = crate::number::try_to_length(ctx, value)?;
    usize::try_from(length).map_err(|_| {
        crate::error::new_error(
            ctx,
            "RangeError",
            "apply argument list exceeds addressable storage",
        )
    })
}

fn reserve_apply_list(ctx: &HostContext, length: usize) -> Result<Vec<Value>, Value> {
    let mut values = Vec::new();
    values.try_reserve_exact(length).map_err(|_| {
        crate::error::new_error(ctx, "RangeError", "Cannot allocate apply argument list")
    })?;
    Ok(values)
}

pub(crate) fn collect_apply_args_inline<const N: usize>(
    value: Option<&Value>,
    inline: &mut [Value; N],
) -> Option<usize> {
    let Some(Value::Object(obj)) = value else {
        return Some(0);
    };

    let object = obj.lock().unwrap();
    if let ObjectKind::Array(values) = &object.kind {
        if values.len() > inline.len() {
            return None;
        }
        for (index, value) in values.iter().enumerate() {
            inline[index] = value.clone();
        }
        return Some(values.len());
    }

    let length = apply_array_like_length(&object);
    if length > inline.len() {
        return None;
    }
    for (index, slot) in inline.iter_mut().enumerate().take(length) {
        *slot = crate::keys::with_index_key(index, |key| {
            object
                .properties
                .get(key)
                .cloned()
                .unwrap_or(Value::Undefined)
        });
    }
    Some(length)
}

pub(crate) fn collect_apply_args(value: &Value) -> Vec<Value> {
    let Value::Object(obj) = value else {
        return Vec::new();
    };

    let object = obj.lock().unwrap();
    if let ObjectKind::Array(values) = &object.kind {
        return values.clone();
    }

    let length = apply_array_like_length(&object);
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

/// Like `bind_function` but reads `__fn_arity` as a length fallback for
/// magic fn_obj mocks (tests pass `{__fn_arity: n}` instead of `length`).
#[cfg(test)]
mod bound_name_allocation_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    fn previous_name(object: &Object) -> Arc<str> {
        let name = match object.properties.get("name").or_else(|| object.properties.get("__fn_name")) {
            Some(Value::String(text)) => text.to_string(),
            Some(other) => crate::keys::value_display_string(other),
            None => String::new(),
        };
        crate::keys::concat2_arc("bound ", &name)
    }

    #[test]
    fn names_preserve_fallbacks_and_current_values() {
        let mut object = Object::new();
        assert_eq!(bound_function_name(&object), previous_name(&object));
        for value in [crate::keys::string_value("named"), crate::keys::string_value(""),
            crate::keys::string_value("bound prior"), crate::keys::string_value("\u{e9}\u{1f680}\0"),
            Value::I32(42), Value::Null, Value::Undefined] {
            object.properties.insert("name".into(), value);
            assert_eq!(bound_function_name(&object), previous_name(&object));
        }
        object.properties.shift_remove("name");
        object.properties.insert("__fn_name".into(), crate::keys::string_value("fallback"));
        assert_eq!(bound_function_name(&object).as_ref(), "bound fallback");
        object.properties.insert("name".into(), crate::keys::string_value("current"));
        assert_eq!(bound_function_name(&object).as_ref(), "bound current");
    }

    #[test]
    fn both_constructors_keep_names_lengths_and_captures() {
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_free_fn("ecma:bind-name-test", "target", Box::new(|_, _| Value::Undefined));
        let index = vm.resolve_host_function_index("ecma:bind-name-test", "target").unwrap();
        let bound_index = vm.resolve_host_function_index("ecma:function", "invokeBound").unwrap();
        let mut target = Object::new();
        target.kind = ObjectKind::HostFunction(index);
        target.properties.insert("name".into(), crate::keys::string_value("target"));
        target.properties.insert("length".into(), Value::I32(3));
        let target = Value::Object(vybe_runtime::heap::alloc(target));
        for constructor in [bind_function, bind_function_with_arity] {
            let result = constructor(&target, &[Value::Null, Value::I32(7)], bound_index);
            let Value::Object(result) = result else { panic!("bound function"); };
            let result = result.lock().unwrap();
            assert!(matches!(result.properties.get("name"), Some(Value::String(name)) if name.as_ref() == "bound target"));
            assert!(matches!(result.properties.get("length"), Some(Value::F64(length)) if *length == 2.0));
            let Some(Value::Object(captures)) = result.properties.get("__bound_args") else { panic!("captures"); };
            let captures = captures.lock().unwrap();
            let ObjectKind::Array(values) = &captures.kind else { panic!("capture array"); };
            assert_eq!(values.len(), 4);
            assert_eq!(values[1], Value::Null);
            assert_eq!(values[3], Value::I32(7));
        }
    }

    #[test]
    #[ignore = "paired bound-name allocation benchmark, not whole binding or guest timing"]
    fn paired_bound_name_benchmark() {
        for length in [0, 8, 64, 512, 4096] {
            let mut object = Object::new();
            object.properties.insert("name".into(), Value::String(Arc::from("x".repeat(length))));
            let timed = |current: bool| {
                let start = Instant::now();
                for _ in 0..10_000 {
                    black_box(if current { bound_function_name(black_box(&object)) } else { previous_name(black_box(&object)) });
                }
                start.elapsed()
            };
            let mut previous = Duration::ZERO;
            let mut current = Duration::ZERO;
            for round in 0..8 {
                if round % 2 == 0 { previous += timed(false); current += timed(true); }
                else { current += timed(true); previous += timed(false); }
            }
            eprintln!("bound name bytes={length}: previous={previous:?}, current={current:?}");
        }
    }
}

fn bound_function_name(object: &Object) -> Arc<str> {
    match object.properties.get("name").or_else(|| object.properties.get("__fn_name")) {
        Some(Value::String(text)) => crate::keys::concat2_arc("bound ", text),
        Some(other) => crate::keys::concat2_arc("bound ", &crate::keys::value_display_string(other)),
        None => crate::keys::string_arc("bound "),
    }
}

fn bind_function_with_arity(target: &Value, bound: &[Value], invoke_bound_idx: usize) -> Value {
    let Value::Object(obj) = target else {
        return target.clone();
    };

    let (
        target_kind,
        existing_bound,
        target_name,
        target_length,
        target_proto,
        target_proto_link,
        target_non_ctor,
        target_proxy_callable,
    ) = {
        let o = obj.lock().unwrap();
        let prev_bound = match o.properties.get("__bound_args") {
            Some(Value::Object(ba)) => {
                let bo = ba.lock().unwrap();
                if let ObjectKind::Array(ref values) = bo.kind {
                    values.clone()
                } else {
                    Vec::new()
                }
            }
            _ => Vec::new(),
        };
        let name = bound_function_name(&o);
        let length = match o.properties.get("length") {
            Some(Value::I32(value)) if *value > 0 => *value as usize,
            Some(Value::I64(value)) if *value > 0 => *value as usize,
            Some(Value::F64(value)) if *value > 0.0 => *value as usize,
            _ => match o.properties.get("__fn_arity") {
                Some(Value::I32(value)) if *value > 0 => *value as usize,
                Some(Value::I64(value)) if *value > 0 => *value as usize,
                Some(Value::F64(value)) if *value > 0.0 => *value as usize,
                _ => 0,
            },
        };
        let prototype = o
            .properties
            .get("prototype")
            .cloned()
            .unwrap_or(Value::Undefined);
        let proto_link = o.properties.get("__proto__").cloned();
        let non_ctor = matches!(o.properties.get("__vybe_non_ctor"), Some(Value::Bool(true)));
        let proxy_callable = matches!(
            o.properties.get("__vybe_proxy_callable"),
            Some(Value::Bool(true))
        );
        (
            o.kind.clone(),
            prev_bound,
            name,
            length,
            prototype,
            proto_link,
            non_ctor,
            proxy_callable,
        )
    };

    if target_proxy_callable && existing_bound.len() >= 2 {
        let mut stored_bound = Vec::with_capacity(2 + bound.len().saturating_sub(1));
        stored_bound.push(existing_bound[0].clone());
        stored_bound.push(bound.first().cloned().unwrap_or(Value::Undefined));
        stored_bound.extend(bound.iter().skip(1).cloned());
        let consumed_args = stored_bound.len().saturating_sub(2);

        let mut wrapper = Object::new();
        wrapper
            .properties
            .reserve(if matches!(target_proto, Value::Null | Value::Undefined) {
                5
            } else {
                6
            });
        wrapper.kind = target_kind;
        wrapper.properties.insert(
            "__bound_args".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(stored_bound))),
        );
        wrapper
            .properties
            .insert("__vybe_proxy_callable".into(), Value::Bool(true));
        wrapper.properties.insert(
            "__proto__".into(),
            match target_proto_link {
                Some(link) if !matches!(link, Value::Null | Value::Undefined) => link,
                _ => shared_function_prototype(),
            },
        );
        wrapper.properties.insert(
            "name".into(),
            Value::String(target_name),
        );
        wrapper.properties.insert(
            "length".into(),
            Value::F64(target_length.saturating_sub(consumed_args) as f64),
        );
        if !matches!(target_proto, Value::Null | Value::Undefined) {
            wrapper.properties.insert("prototype".into(), target_proto);
        }
        return Value::Object(vybe_runtime::heap::alloc(wrapper));
    }

    // Allow ordinary objects (magic fn_obj descriptors from tests) — don't bail for non-Function.
    let is_existing_bound_wrapper = matches!(&target_kind, ObjectKind::HostFunction(idx) if *idx == invoke_bound_idx)
        && existing_bound.len() >= 3;
    let mut stored_bound = if is_existing_bound_wrapper {
        Vec::with_capacity(existing_bound.len() + bound.len().saturating_sub(1))
    } else {
        Vec::with_capacity(3 + bound.len().saturating_sub(1))
    };
    if is_existing_bound_wrapper {
        stored_bound.push(existing_bound[0].clone());
        stored_bound.push(existing_bound[1].clone());
        stored_bound.push(existing_bound[2].clone());
        stored_bound.extend(existing_bound.iter().skip(3).cloned());
        stored_bound.extend(bound.iter().skip(1).cloned());
    } else {
        stored_bound.push(target.clone());
        stored_bound.push(bound.first().cloned().unwrap_or(Value::Undefined));
        stored_bound.push(target_proto.clone());
        stored_bound.extend(bound.iter().skip(1).cloned());
    }

    let mut wrapper = Object::new();
    wrapper.properties.reserve(
        4 + usize::from(target_non_ctor)
            + usize::from(!matches!(target_proto, Value::Null | Value::Undefined)),
    );
    wrapper.kind = ObjectKind::HostFunction(invoke_bound_idx);
    let consumed_args = stored_bound.len().saturating_sub(3);
    wrapper.properties.insert(
        "__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(stored_bound))),
    );
    // §10.4.1.3 BoundFunctionCreate step 1: the bound function's
    // [[Prototype]] is the TARGET's [[Prototype]] — a bound async fn
    // stays `instanceof AsyncFunction`, a bound generator fn stays
    // `instanceof GeneratorFunction`. Fall back to %Function.prototype%.
    wrapper.properties.insert(
        "__proto__".into(),
        match target_proto_link {
            Some(link) if !matches!(link, Value::Null | Value::Undefined) => link,
            _ => shared_function_prototype(),
        },
    );
    if target_non_ctor {
        wrapper
            .properties
            .insert("__vybe_non_ctor".into(), Value::Bool(true));
    }
    wrapper.properties.insert(
        "name".into(),
        Value::String(target_name),
    );
    wrapper.properties.insert(
        "length".into(),
        Value::F64(target_length.saturating_sub(consumed_args) as f64),
    );
    if !matches!(target_proto, Value::Null | Value::Undefined) {
        wrapper.properties.insert("prototype".into(), target_proto);
    }
    Value::Object(vybe_runtime::heap::alloc(wrapper))
}

/// Build a function ref carrying bound args. Mirrors the convention in
/// `crate::bound_host_fn_ref` but works on any function-like
/// Value (HostFunction or user Function).
fn bind_function(target: &Value, bound: &[Value], invoke_bound_idx: usize) -> Value {
    let Value::Object(obj) = target else {
        return target.clone();
    };

    let (
        target_kind,
        existing_bound,
        target_name,
        target_length,
        target_proto,
        target_proto_link,
        target_non_ctor,
        target_proxy_callable,
    ) = {
        let o = obj.lock().unwrap();
        let prev_bound = match o.properties.get("__bound_args") {
            Some(Value::Object(ba)) => {
                let bo = ba.lock().unwrap();
                if let ObjectKind::Array(ref values) = bo.kind {
                    values.clone()
                } else {
                    Vec::new()
                }
            }
            _ => Vec::new(),
        };
        let name = bound_function_name(&o);
        let length = match o.properties.get("length") {
            Some(Value::I32(value)) if *value > 0 => *value as usize,
            Some(Value::I64(value)) if *value > 0 => *value as usize,
            Some(Value::F64(value)) if *value > 0.0 => *value as usize,
            _ => match o.properties.get("__fn_arity") {
                Some(Value::I32(value)) if *value > 0 => *value as usize,
                Some(Value::I64(value)) if *value > 0 => *value as usize,
                Some(Value::F64(value)) if *value > 0.0 => *value as usize,
                _ => 0,
            },
        };
        let prototype = o
            .properties
            .get("prototype")
            .cloned()
            .unwrap_or(Value::Undefined);
        let proto_link = o.properties.get("__proto__").cloned();
        let non_ctor = matches!(o.properties.get("__vybe_non_ctor"), Some(Value::Bool(true)));
        let proxy_callable = matches!(
            o.properties.get("__vybe_proxy_callable"),
            Some(Value::Bool(true))
        );
        (
            o.kind.clone(),
            prev_bound,
            name,
            length,
            prototype,
            proto_link,
            non_ctor,
            proxy_callable,
        )
    };

    if target_proxy_callable && existing_bound.len() >= 2 {
        let mut stored_bound = Vec::with_capacity(2 + bound.len().saturating_sub(1));
        stored_bound.push(existing_bound[0].clone());
        stored_bound.push(bound.first().cloned().unwrap_or(Value::Undefined));
        stored_bound.extend(bound.iter().skip(1).cloned());
        let consumed_args = stored_bound.len().saturating_sub(2);

        let mut wrapper = Object::new();
        wrapper
            .properties
            .reserve(if matches!(target_proto, Value::Null | Value::Undefined) {
                5
            } else {
                6
            });
        wrapper.kind = target_kind;
        wrapper.properties.insert(
            "__bound_args".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(stored_bound))),
        );
        wrapper
            .properties
            .insert("__vybe_proxy_callable".into(), Value::Bool(true));
        wrapper.properties.insert(
            "__proto__".into(),
            match target_proto_link {
                Some(link) if !matches!(link, Value::Null | Value::Undefined) => link,
                _ => shared_function_prototype(),
            },
        );
        wrapper.properties.insert(
            "name".into(),
            Value::String(target_name),
        );
        wrapper.properties.insert(
            "length".into(),
            Value::F64(target_length.saturating_sub(consumed_args) as f64),
        );
        if !matches!(target_proto, Value::Null | Value::Undefined) {
            wrapper.properties.insert("prototype".into(), target_proto);
        }
        return Value::Object(vybe_runtime::heap::alloc(wrapper));
    }

    // Allow ordinary objects (magic fn_obj descriptors from tests) — don't bail for non-Function.
    let is_existing_bound_wrapper = matches!(&target_kind, ObjectKind::HostFunction(idx) if *idx == invoke_bound_idx)
        && existing_bound.len() >= 3;
    let mut stored_bound = if is_existing_bound_wrapper {
        Vec::with_capacity(existing_bound.len() + bound.len().saturating_sub(1))
    } else {
        Vec::with_capacity(3 + bound.len().saturating_sub(1))
    };
    if is_existing_bound_wrapper {
        stored_bound.push(existing_bound[0].clone());
        stored_bound.push(existing_bound[1].clone());
        stored_bound.push(existing_bound[2].clone());
        stored_bound.extend(existing_bound.iter().skip(3).cloned());
        stored_bound.extend(bound.iter().skip(1).cloned());
    } else {
        stored_bound.push(target.clone());
        stored_bound.push(bound.first().cloned().unwrap_or(Value::Undefined));
        stored_bound.push(target_proto.clone());
        stored_bound.extend(bound.iter().skip(1).cloned());
    }

    let mut wrapper = Object::new();
    wrapper.properties.reserve(
        4 + usize::from(target_non_ctor)
            + usize::from(!matches!(target_proto, Value::Null | Value::Undefined)),
    );
    wrapper.kind = ObjectKind::HostFunction(invoke_bound_idx);
    let consumed_args = stored_bound.len().saturating_sub(3);
    wrapper.properties.insert(
        "__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(stored_bound))),
    );
    // §10.4.1.3 BoundFunctionCreate step 1: the bound function's
    // [[Prototype]] is the TARGET's [[Prototype]] — a bound async fn
    // stays `instanceof AsyncFunction`, a bound generator fn stays
    // `instanceof GeneratorFunction`. Fall back to %Function.prototype%.
    wrapper.properties.insert(
        "__proto__".into(),
        match target_proto_link {
            Some(link) if !matches!(link, Value::Null | Value::Undefined) => link,
            _ => shared_function_prototype(),
        },
    );
    if target_non_ctor {
        wrapper
            .properties
            .insert("__vybe_non_ctor".into(), Value::Bool(true));
    }
    wrapper.properties.insert(
        "name".into(),
        Value::String(target_name),
    );
    wrapper.properties.insert(
        "length".into(),
        Value::F64(target_length.saturating_sub(consumed_args) as f64),
    );
    if !matches!(target_proto, Value::Null | Value::Undefined) {
        wrapper.properties.insert("prototype".into(), target_proto);
    }
    Value::Object(vybe_runtime::heap::alloc(wrapper))
}

#[cfg(test)]
mod initializer_buffer_speedup_tests {
    use super::*;

    fn original_prepare(list: Option<&Value>, arity: usize, receiver: &Value) -> Vec<Value> {
        let mut first: [Value; APPLY_INLINE_ARG_LIMIT] = std::array::from_fn(|_| Value::Undefined);
        if let Some(length) = collect_apply_args_inline(list, &mut first) {
            if arity <= APPLY_INLINE_ARG_LIMIT {
                let mut second: [Value; APPLY_INLINE_ARG_LIMIT + 1] = std::array::from_fn(|_| Value::Undefined);
                for index in 0..length.min(arity) { second[index] = first[index].clone(); }
                second[arity] = receiver.clone();
                return second[..arity + 1].to_vec();
            }
        }
        let mut packed = list.map(collect_apply_args).unwrap_or_default();
        packed.resize(arity, Value::Undefined);
        packed.push(receiver.clone());
        packed
    }

    fn array(length: usize) -> Value {
        Value::Object(vybe_runtime::heap::alloc(Object::new_array((0..length).map(|i| Value::I32(i as i32)).collect())))
    }

    fn target(vm: &VM, module: &str, name: &str) -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    #[test]
    fn single_buffer_preserves_arity_padding_truncation_and_receiver_identity() {
        let receiver = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        for arity in [0, 1, 4, 8, 9, 16] {
            for count in [0, 1, 4, 8, 9, 10, 16] {
                let list = array(count);
                let expected = original_prepare(Some(&list), arity, &receiver);
                let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] = std::array::from_fn(|_| Value::Undefined);
                let actual = if let Some(length) = collect_initializer_inline(Some(&list), arity, &mut inline) {
                    inline[arity] = receiver.clone();
                    inline[..length].to_vec()
                } else {
                    let mut packed = collect_apply_args(&list);
                    packed.resize(arity, Value::Undefined);
                    packed.push(receiver.clone());
                    packed
                };
                assert_eq!(actual.len(), expected.len());
                for (actual, expected) in actual.iter().zip(&expected) {
                    match (actual, expected) {
                        (Value::I32(a), Value::I32(b)) => assert_eq!(a, b),
                        (Value::Undefined, Value::Undefined) => {},
                        (Value::Object(a), Value::Object(b)) => assert!(Arc::ptr_eq(a, b)),
                        _ => panic!("initializer argument mismatch"),
                    }
                }
            }
        }
    }

    #[test]
    fn registered_initialization_reads_current_metadata_and_releases_class_lock() {
        let mut vm = VM::new();
        register(&mut vm);
        let class = vybe_runtime::heap::alloc(Object::new());
        let class_for_callback = class.clone();
        let seen = Arc::new(Mutex::new(Vec::<Value>::new()));
        let recorded = seen.clone();
        vm.register_free_fn("test:initializer", "capture", Box::new(move |_, args| {
            assert!(class_for_callback.try_lock().is_ok());
            *recorded.lock().unwrap() = args.to_vec();
            Value::I32(321)
        }));
        let initializer = target(&vm, "test:initializer", "capture");
        class.lock().unwrap().properties.insert("__vybe_constructor_initializer".into(), initializer);
        let initialize = target(&vm, "ecma:function", "initialize");
        let receiver = vybe_runtime::heap::alloc(Object::new());
        let args = [Value::Object(class.clone()), Value::Object(receiver.clone()), array(9)];
        for arity in [8, 0, 16] {
            class.lock().unwrap().properties.insert("__vybe_constructor_arity".into(), Value::I32(arity));
            assert!(matches!(vm.invoke(&initialize, &args).unwrap(), Value::I32(321)));
            let seen = seen.lock().unwrap();
            assert_eq!(seen.len(), arity as usize + 1);
            for index in 0..arity as usize {
                if index < 9 { assert!(matches!(seen[index], Value::I32(value) if value == index as i32)); }
                else { assert!(matches!(seen[index], Value::Undefined)); }
            }
            let Value::Object(actual) = &seen[arity as usize] else { panic!("receiver missing") };
            assert!(Arc::ptr_eq(actual, &receiver));
        }
    }

    #[test]
    fn sparse_and_generic_lists_match_original_at_inline_boundary() {
        let receiver = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let check = |list: &Value| {
            for arity in [0, 1, 4, 8, 9, 16] {
                let expected = original_prepare(Some(list), arity, &receiver);
                let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] = std::array::from_fn(|_| Value::Undefined);
                let actual = if let Some(length) = collect_initializer_inline(Some(list), arity, &mut inline) {
                    inline[arity] = receiver.clone();
                    inline[..length].to_vec()
                } else {
                    let mut packed = collect_apply_args(list);
                    packed.resize(arity, Value::Undefined);
                    packed.push(receiver.clone());
                    packed
                };
                assert_eq!(actual.len(), expected.len());
                for (index, (actual, expected)) in actual.iter().zip(&expected).enumerate() {
                    match (actual, expected) {
                        (Value::I32(a), Value::I32(b)) => assert_eq!(a, b, "arity={arity}, index={index}"),
                        (Value::Undefined, Value::Undefined) => {},
                        (Value::Object(a), Value::Object(b)) => assert!(Arc::ptr_eq(a, b)),
                        _ => panic!("array-like initializer mismatch at arity={arity}, index={index}"),
                    }
                }
            }
        };
        for length in [0, 1, 8, 9, 10, 16] {
            let mut object = Object::new();
            object.properties.insert("length".into(), Value::F64(length as f64));
            object.properties.insert("unrelated".into(), Value::I32(999));
            // Reverse insertion order and holes must not change argument order.
            for index in (0..length).rev() {
                if index % 3 != 1 { object.properties.insert(index.to_string(), Value::I32(index as i32)); }
            }
            check(&Value::Object(vybe_runtime::heap::alloc(object)));
            let values = (0..length).map(|index| if index % 3 == 1 {
                Value::Undefined
            } else { Value::I32(index as i32) }).collect();
            check(&Value::Object(vybe_runtime::heap::alloc(Object::new_array(values))));
        }
        for length in [Value::I32(-1), Value::F64(9.5), Value::Null, Value::Undefined] {
            let mut object = Object::new();
            object.properties.insert("length".into(), length);
            object.properties.insert("0".into(), Value::I32(42));
            object.properties.insert("8".into(), Value::I32(84));
            check(&Value::Object(vybe_runtime::heap::alloc(object)));
        }
        for list in [Value::Null, Value::Undefined, Value::I32(7)] { check(&list); }
    }

    #[test]
    #[ignore = "native initializer packing benchmark, not constructor/guest execution timing"]
    fn initializer_packing_benchmark() {
        use std::{hint::black_box, time::Instant};
        let list = array(4);
        let receiver = Value::Undefined;
        let start = Instant::now();
        for _ in 0..10000 {
            let mut first: [Value; APPLY_INLINE_ARG_LIMIT] = std::array::from_fn(|_| Value::Undefined);
            let length = collect_apply_args_inline(Some(black_box(&list)), &mut first).unwrap();
            let mut second: [Value; APPLY_INLINE_ARG_LIMIT + 1] = std::array::from_fn(|_| Value::Undefined);
            for index in 0..length.min(4) { second[index] = first[index].clone(); }
            second[4] = receiver.clone();
            black_box(&second[..5]);
        }
        let original = start.elapsed();
        let start = Instant::now();
        for _ in 0..10000 {
            let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] = std::array::from_fn(|_| Value::Undefined);
            let length = collect_initializer_inline(Some(black_box(&list)), 4, &mut inline).unwrap();
            inline[4] = receiver.clone();
            black_box(&inline[..length]);
        }
        eprintln!("initializer 10000 packs: two-buffer={original:?}, single-buffer={:?}", start.elapsed());
    }

    #[test]
    #[ignore = "registered initializer benchmark with native stub, not guest constructor timing"]
    fn registered_initializer_benchmark() {
        use std::{hint::black_box, time::Instant};
        let mut vm = VM::new();
        register(&mut vm);
        vm.register_free_fn("test:initializer", "stub", Box::new(|_, _| Value::I32(123)));
        vm.register_host_fn("test:initializer", "original", Box::new(|ctx, args| {
            let receiver = args.get(1).cloned().unwrap_or(Value::Undefined);
            let Some(Value::Object(class)) = args.first() else { return receiver };
            let metadata = {
                let class = class.lock().unwrap();
                class.properties.get("__vybe_constructor_initializer").cloned().zip(
                    class.properties.get("__vybe_constructor_arity").map(|value| value.as_f64() as usize)
                )
            };
            let Some((initializer, arity)) = metadata else { return receiver };
            let mut first: [Value; APPLY_INLINE_ARG_LIMIT] = std::array::from_fn(|_| Value::Undefined);
            if let Some(length) = collect_apply_args_inline(args.get(2), &mut first) {
                if arity <= APPLY_INLINE_ARG_LIMIT {
                    let mut second: [Value; APPLY_INLINE_ARG_LIMIT + 1] = std::array::from_fn(|_| Value::Undefined);
                    for index in 0..length.min(arity) { second[index] = first[index].clone(); }
                    second[arity] = receiver;
                    return ctx.invoke(&initializer, &second[..arity + 1]);
                }
            }
            let mut packed = args.get(2).map(collect_apply_args).unwrap_or_default();
            packed.resize(arity, Value::Undefined);
            packed.push(receiver);
            ctx.invoke(&initializer, &packed)
        }));
        let mut class = Object::new();
        class.properties.insert("__vybe_constructor_initializer".into(), target(&vm, "test:initializer", "stub"));
        class.properties.insert("__vybe_constructor_arity".into(), Value::I32(4));
        let args = [Value::Object(vybe_runtime::heap::alloc(class)),
            Value::Object(vybe_runtime::heap::alloc(Object::new())), array(4)];
        let original = target(&vm, "test:initializer", "original");
        let optimized = target(&vm, "ecma:function", "initialize");
        assert!(matches!(vm.invoke(&original, &args).unwrap(), Value::I32(123)));
        assert!(matches!(vm.invoke(&optimized, &args).unwrap(), Value::I32(123)));
        // Alternate paired batch order to reduce first/second scheduling bias.
        let mut original_time = std::time::Duration::ZERO;
        let mut optimized_time = std::time::Duration::ZERO;
        for round in 0..6 {
            let order = if round % 2 == 0 { [(&original, false), (&optimized, true)] }
                else { [(&optimized, true), (&original, false)] };
            for (function, is_optimized) in order {
                let start = Instant::now();
                for _ in 0..2000 { black_box(vm.invoke(function, black_box(&args)).unwrap()); }
                if is_optimized { optimized_time += start.elapsed(); }
                else { original_time += start.elapsed(); }
            }
        }
        eprintln!("initializer 12000 registered calls: original={original_time:?}, optimized={optimized_time:?}");
    }
}

#[cfg(test)]
mod apply_target_classification_tests {
    use super::*;

    fn guest() -> Value {
        let mut object = Object::new();
        object.kind = ObjectKind::Function(vybe_runtime::value::Function {
            name: Some("apply-classification".into()),
            arity: 1,
            chunk_index: 0,
            upvalues: Vec::new(),
        });
        Value::Object(vybe_runtime::heap::alloc(object))
    }

    fn legacy(target: &Value) -> (bool, bool) {
        let host = is_host_callee(target);
        (host || is_callable(target), host)
    }

    #[test]
    fn classification_matches_existing_callable_and_proxy_rules() {
        let guest = guest();
        let mut host = Object::new();
        host.kind = ObjectKind::HostFunction(0);
        let host = Value::Object(vybe_runtime::heap::alloc(host));
        let mut proxy = Object::new();
        proxy.kind = ObjectKind::HostFunction(0);
        proxy.properties.insert("__vybe_proxy_target".into(), guest.clone());
        proxy.properties.insert("__vybe_proxy_handler".into(), Value::Null);
        let proxy = Value::Object(vybe_runtime::heap::alloc(proxy));
        let ordinary = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        for (target, expected) in [
            (guest, (true, false)),
            (host, (true, true)),
            (proxy, (true, false)),
            (ordinary, (false, false)),
            (Value::Null, (false, false)),
            (Value::Undefined, (false, false)),
            (Value::I32(1), (false, false)),
        ] {
            assert_eq!(legacy(&target), expected);
            assert_eq!(apply_target_kind(&target), expected);
        }
    }

    #[test]
    #[ignore = "paired target classification microbenchmark, not guest execution timing"]
    fn paired_classification_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        let guest = guest();
        let mut host = Object::new();
        host.kind = ObjectKind::HostFunction(0);
        let host = Value::Object(vybe_runtime::heap::alloc(host));
        for (name, target) in [("guest", guest), ("host", host)] {
            let mut old = Duration::ZERO;
            let mut new = Duration::ZERO;
            for round in 0..8 {
                for optimized in if round % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..100_000 {
                        let target = black_box(&target);
                        black_box(if optimized { apply_target_kind(target) } else { legacy(target) });
                    }
                    if optimized { new += start.elapsed(); } else { old += start.elapsed(); }
                }
            }
            eprintln!("apply {name} classification 800000 calls: previous={old:?}, fused={new:?}");
        }
    }
}

#[cfg(test)]
mod rest_array_bulk_clone_candidate_tests {
    use super::*;

    fn candidate(args: &[Value], fixed_count: usize) -> Value {
        let rest = args[fixed_count.min(args.len())..].to_vec();
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(rest)))
    }

    fn contents(value: &Value) -> Vec<Value> {
        let Value::Object(object) = value else { panic!("expected rest array") };
        let object = object.lock().unwrap();
        let ObjectKind::Array(values) = &object.kind else { panic!("expected array backing") };
        values.clone()
    }

    #[test]
    fn bulk_clone_keeps_fresh_backing_and_shared_element_identity() {
        let shared = vybe_runtime::heap::alloc(Object::new());
        let text: Arc<str> = Arc::from("rest");
        let args = [Value::I32(7), Value::Object(shared.clone()), Value::String(text.clone())];
        for fixed in [0, 1, 2, 3, 8, usize::MAX] {
            let old = make_rest_array(&args, fixed);
            let new = candidate(&args, fixed);
            assert!(matches!((&old, &new), (Value::Object(a), Value::Object(b)) if !Arc::ptr_eq(a, b)));
            let old_values = contents(&old);
            let new_values = contents(&new);
            assert_eq!(new_values.len(), args.len().saturating_sub(fixed));
            assert_eq!(old_values.len(), new_values.len());
            for (old, new) in old_values.iter().zip(&new_values) {
                match (old, new) {
                    (Value::I32(a), Value::I32(b)) => assert_eq!(a, b),
                    (Value::Object(a), Value::Object(b)) => assert!(Arc::ptr_eq(a, b)),
                    (Value::String(a), Value::String(b)) => assert!(Arc::ptr_eq(a, b)),
                    _ => panic!("rest element shape changed"),
                }
            }
        }
        let first = candidate(&args, 0);
        let second = candidate(&args, 0);
        let Value::Object(first) = first else { unreachable!() };
        if let ObjectKind::Array(values) = &mut first.lock().unwrap().kind {
            values[0] = Value::I32(99);
        }
        assert!(matches!(contents(&second)[0], Value::I32(7)));
        assert!(matches!(args[0], Value::I32(7)));
    }

    #[test]
    #[ignore = "paired rest-array allocation benchmark, not whole callback timing"]
    fn paired_rest_array_allocation_benchmark() {
        use std::{hint::black_box, time::{Duration, Instant}};
        let shared = Value::Object(vybe_runtime::heap::alloc(Object::new()));
        let text = crate::keys::string_value("rest-benchmark");
        for length in [0, 1, 8, 64, 512] {
            let mut args = vec![Value::I32(0); length + 2];
            for (index, value) in args.iter_mut().enumerate() {
                *value = match index % 3 {
                    0 => Value::I32(index as i32),
                    1 => shared.clone(),
                    _ => text.clone(),
                };
            }
            let mut previous = Duration::ZERO;
            let mut bulk = Duration::ZERO;
            for round in 0..8 {
                for optimized in if round % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..1000 {
                        black_box(if optimized { candidate(black_box(&args), 2) }
                            else { make_rest_array(black_box(&args), 2) });
                    }
                    if optimized { bulk += start.elapsed(); } else { previous += start.elapsed(); }
                }
            }
            eprintln!("rest array len={length} 8000 allocations: previous={previous:?}, bulk={bulk:?}");
        }
    }
}
