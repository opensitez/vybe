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
            let mut inline_args: [Value; APPLY_INLINE_ARG_LIMIT] =
                std::array::from_fn(|_| Value::Undefined);
            if let Some(inline_len) = collect_apply_args_inline(args.get(2), &mut inline_args) {
                if arity <= APPLY_INLINE_ARG_LIMIT {
                    let mut inline: [Value; APPLY_INLINE_ARG_LIMIT + 1] =
                        std::array::from_fn(|_| Value::Undefined);
                    for index in 0..inline_len.min(arity) {
                        inline[index] = inline_args[index].clone();
                    }
                    inline[arity] = receiver;
                    return ctx.invoke(&initializer, &inline[..arity + 1]);
                }
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
            let host_callee = is_host_callee(target);
            if !host_callee && !is_callable(target) {
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
        for value in &prepared.partials {
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
        for value in &prepared.partials {
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
        ExplicitThisCallKind::Compiled => {
            // Arrows need no special case here: they capture lexical
            // `this` at creation (compiler-emitted upvalue) and never read
            // the ambient binding this sets (§10.2.11).
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = invoke_compiled_function(ctx, target, args);
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
        ExplicitThisCallKind::Compiled => {
            let previous_this = ctx.current_js_this();
            ctx.set_js_this(this_arg);
            let result = try_invoke_compiled_function(ctx, target, args);
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
    let Some(fixed_count) = compiled_rest_fixed_arity(target) else {
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
    let Some(fixed_count) = compiled_rest_fixed_arity(target) else {
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
    Compiled,
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
        ObjectKind::Function(_) => ExplicitThisCallKind::Compiled,
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
        let name = match o
            .properties
            .get("name")
            .or_else(|| o.properties.get("__fn_name"))
        {
            Some(Value::String(text)) => text.to_string(),
            Some(other) => crate::keys::value_display_string(other),
            None => String::new(),
        };
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
            Value::String(crate::keys::concat2_arc("bound ", &target_name)),
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
        Value::String(crate::keys::concat2_arc("bound ", &target_name)),
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
        let name = match o
            .properties
            .get("name")
            .or_else(|| o.properties.get("__fn_name"))
        {
            Some(Value::String(text)) => text.to_string(),
            Some(other) => crate::keys::value_display_string(other),
            None => String::new(),
        };
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
            Value::String(crate::keys::concat2_arc("bound ", &target_name)),
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
        Value::String(crate::keys::concat2_arc("bound ", &target_name)),
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
