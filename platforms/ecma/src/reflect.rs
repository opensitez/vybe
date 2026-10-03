//! ECMA-262 §28.1 — Reflect.
//!
//! Static methods that mirror the corresponding object-operation
//! abstract ops in the spec:
//!
//!   §28.1.1 Reflect.apply(target, thisArg, argsList)
//!   §28.1.2 Reflect.construct(target, argsList, newTarget?)
//!   §28.1.3 Reflect.defineProperty(target, key, attrs)
//!   §28.1.4 Reflect.deleteProperty(target, key)
//!   §28.1.5 Reflect.get(target, key, receiver?)
//!   §28.1.6 Reflect.getOwnPropertyDescriptor(target, key)
//!   §28.1.7 Reflect.getPrototypeOf(target)
//!   §28.1.8 Reflect.has(target, key)
//!   §28.1.9 Reflect.isExtensible(target)
//!   §28.1.10 Reflect.ownKeys(target)
//!   §28.1.11 Reflect.preventExtensions(target)
//!   §28.1.12 Reflect.set(target, key, value, receiver?)
//!   §28.1.13 Reflect.setPrototypeOf(target, proto)
//!
//! Most are thin forwards to `ecma:object.*` because the underlying
//! Object operations are the same — Reflect just exposes them as
//! standalone functions instead of Object statics.

use crate::function::{collect_apply_args_once, invoke_with_explicit_this};
use crate::object::{
    install_noop_setter, is_nonconfig, is_not_extensible, mark_not_extensible,
    ordered_own_string_keys, proto_walk_has, proxy_target_and_handler, proxy_trap, track_key,
    track_nonconfig, track_nonenum,
};
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

#[inline]
fn proxy_revoke_host_key() -> &'static (String, String) {
    static KEY: std::sync::OnceLock<(String, String)> = std::sync::OnceLock::new();
    KEY.get_or_init(|| ("ecma:reflect".to_string(), "__proxyRevoke".to_string()))
}

// §28.1 step 1 of most Reflect ops: "If target is not an Object, throw a
// TypeError". Thrown as a real error object so `e instanceof TypeError`
// holds in the catcher.
fn throw_type_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
    Value::Undefined
}

fn numeric_index(key: &str) -> Option<usize> {
    crate::keys::non_negative_integer_index_key(key)
}

fn with_property_key<R>(value: Option<&Value>, f: impl FnOnce(&str) -> R) -> R {
    match value {
        Some(value) => crate::keys::with_property_key(value, f),
        None => f(""),
    }
}

/// §28.1.5 Reflect.get — proxy trap, typed-array elements, and a
/// [[Get]]-shaped prototype walk that invokes accessors with `receiver`.
fn reflect_get(ctx: &mut HostContext, target: &Value, key: &str, receiver: Value) -> Value {
    match try_reflect_get(ctx, target, key, receiver) {
        Ok(value) => value,
        Err(error) => {
            ctx.throw_value(error);
            Value::Undefined
        }
    }
}

pub(crate) fn try_reflect_get(
    ctx: &mut HostContext,
    target: &Value,
    key: &str,
    receiver: Value,
) -> Result<Value, Value> {
    let Value::Object(obj) = target else {
        return Err(crate::error::new_error(
            ctx,
            "TypeError",
            "Reflect.get called on non-object",
        ));
    };
    if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
        if let Some(trap) = proxy_trap(&handler, "get") {
            return crate::function::try_invoke_with_explicit_this(
                ctx,
                &trap,
                handler,
                &[proxy_target, crate::keys::string_value(key), receiver],
            );
        }
        return try_reflect_get(ctx, &proxy_target, key, receiver);
    }
    {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(ref values) = o.kind {
            if key == "length" {
                return Ok(Value::I32(values.len() as i32));
            }
            if let Some(i) = numeric_index(key) {
                let accessor = crate::keys::with_getter_property_key(key, |getter_key| {
                    o.properties.contains_key(getter_key)
                });
                if !accessor && !crate::array::is_array_hole(&o, i) {
                    if let Some(value) = values.get(i) {
                        return Ok(value.clone());
                    }
                }
            }
        }
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            if let Some(i) = numeric_index(key) {
                if i < ta.length {
                    return Ok(crate::typedarray::read_element(ta, i));
                }
                return Ok(Value::Undefined);
            }
        }
    }
    // [[Get]] walk: accessor (getter) beats data at each level; the
    // getter runs with `this = receiver` (§10.1.8.1 step 8).
    crate::keys::with_getter_property_key(key, |getter_key| {
        let mut current = Some(obj.clone());
        while let Some(node) = current {
            let (getter, data, next) =
                {
                    let o = node.lock().unwrap();
                    (
                        o.properties.get(getter_key).cloned(),
                        o.properties.get(key).cloned(),
                        Some(o.properties.get("__proto__").cloned().unwrap_or_else(|| {
                            crate::object::implicit_object_prototype(&node, &o)
                        })),
                    )
                };
            if let Some(g @ Value::Object(_)) = getter {
                // Accessor convention (see proto_walk_invoke_getter): arity-0
                // getters read ambient `this`; arity-1 getters take the
                // receiver as arg 0. Either way `this = receiver` (§10.1.8.1).
                let arity = match &g {
                    Value::Object(go) => match &go.lock().unwrap().kind {
                        ObjectKind::Function(f) => f.arity,
                        _ => 0,
                    },
                    _ => 0,
                };
                if arity == 0 {
                    return crate::function::try_invoke_with_explicit_this(ctx, &g, receiver, &[]);
                }
                return crate::function::try_invoke_with_explicit_this(
                    ctx,
                    &g,
                    receiver.clone(),
                    &[receiver],
                );
            }
            if let Some(v) = data {
                return Ok(v);
            }
            current = match next {
                Some(Value::Object(p)) => Some(p),
                _ => None,
            };
        }
        Ok(Value::Undefined)
    })
}

/// §28.1.12 Reflect.set — proxy trap, typed-array elements, setter
/// dispatch with `receiver`, non-writable/non-extensible gates.
fn reflect_set(
    ctx: &mut HostContext,
    target: &Value,
    key: &str,
    val: Value,
    receiver: Value,
) -> Value {
    let Value::Object(obj) = target else {
        return throw_type_error(ctx, "Reflect.set called on non-object");
    };
    if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
        if let Some(trap) = proxy_trap(&handler, "set") {
            let result = invoke_with_explicit_this(
                ctx,
                &trap,
                handler,
                &[proxy_target, crate::keys::string_value(key), val, receiver],
            );
            return Value::Bool(result.as_bool());
        }
        // Default receiver is the proxy itself — rebind it to the target
        // so the eventual data write lands on the target (the observable
        // §10.5.9 outcome for trap-less proxies).
        let follow_receiver = match (&receiver, target) {
            (Value::Object(r), Value::Object(t)) if Arc::ptr_eq(r, t) => proxy_target.clone(),
            _ => receiver,
        };
        return reflect_set(ctx, &proxy_target, key, val, follow_receiver);
    }
    {
        let o = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            if let Some(i) = numeric_index(key) {
                if i < ta.length {
                    crate::typedarray::write_element(ta, i, &val);
                    return Value::Bool(true);
                }
                return Value::Bool(false);
            }
        }
    }
    {
        let mut o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Array(_)) {
            if key == "length" {
                crate::array::apply_js_array_length(ctx, &mut o, &val);
                return Value::Bool(true);
            }
            if let Some(i) = numeric_index(key) {
                if o.properties.get("__array_length_readonly").is_some() {
                    return Value::Bool(false);
                }
                if let ObjectKind::Array(ref mut values) = o.kind {
                    if i >= values.len() {
                        values.resize(i + 1, Value::Undefined);
                    }
                    values[i] = val;
                }
                drop(o);
                track_key(obj, key);
                return Value::Bool(true);
            }
        }
    }
    // Walk for an accessor: a REAL setter (compiled Function) runs with
    // `this = receiver` (§10.1.9.2); the noop setter installed for
    // non-writable data properties (a HostFunction) means reject. A
    // getter with no setter at the same level also rejects.
    let accessor_result = crate::keys::with_setter_property_key(key, |setter_key| {
        crate::keys::with_getter_property_key(key, |getter_key| {
            let mut current = Some(obj.clone());
            while let Some(node) = current {
                let (setter, has_getter, has_data, next) = {
                    let o = node.lock().unwrap();
                    (
                        o.properties.get(setter_key).cloned(),
                        o.properties.contains_key(getter_key),
                        o.properties.contains_key(key),
                        o.properties.get("__proto__").cloned(),
                    )
                };
                if let Some(Value::Object(s_obj)) = setter {
                    let arity = match &s_obj.lock().unwrap().kind {
                        ObjectKind::Function(f) => Some(f.arity),
                        _ => None,
                    };
                    if let Some(arity) = arity {
                        // Accessor convention: arity-1 setters take (value) with
                        // ambient `this`; arity-2 setters take (receiver, value).
                        let st = Value::Object(s_obj);
                        // ⛔ ONE RECEIVER, PASSED ONCE. The arity guess below was
                        // reading a receiver-first setter as `(receiver, value)` and
                        // passing both — but under `ReceiverAbi::Parameter` the invoke
                        // ALSO supplies the receiver, so the setter got it twice:
                        // `this` took the supplied one and the value parameter took the
                        // receiver. Measured: `Reflect.set(t,"v",5)` wrote `[object]`.
                        // `invoke_with_receiver` is correct under BOTH bindings, so the
                        // arity question disappears rather than being answered.
                        let _ = arity;
                        ctx.invoke_with_receiver(&st, receiver.clone(), &[val.clone()]);
                        return Some(Value::Bool(true));
                    }
                    // noop setter (HostFunction) = non-writable data property
                    return Some(Value::Bool(false));
                }
                if has_getter {
                    // §10.1.9.2 step 6.c: accessor without a [[Set]] → false
                    return Some(Value::Bool(false));
                }
                if has_data {
                    break; // data property found — assign below
                }
                current = match next {
                    Some(Value::Object(p)) => Some(p),
                    _ => None,
                };
            }
            None
        })
    });
    if let Some(value) = accessor_result {
        return value;
    }
    // Ordinary data assignment onto the receiver (defaults to target).
    let dest = match &receiver {
        Value::Object(recv) => recv.clone(),
        _ => obj.clone(),
    };
    {
        let mut o = dest.lock().unwrap();
        if is_not_extensible(&o) && !o.properties.contains_key(key) {
            return Value::Bool(false);
        }
        o.properties.insert(key.to_string(), val);
    }
    track_key(&dest, key);
    Value::Bool(true)
}

pub fn register(vm: &mut VM) {
    crate::perf::register_host_fn(
        vm,
        "ecma:reflect",
        "__proxyRevoke",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(proxy)) = args.first() {
                let mut o = proxy.lock().unwrap();
                o.properties
                    .insert("__vybe_proxy_revoked".into(), Value::Bool(true));
                o.properties
                    .insert("__vybe_proxy_handler".into(), Value::Null);
            }
            Value::Undefined
        }),
    );
    let proxy_revoke_idx = *vm
        .host_registry
        .get(proxy_revoke_host_key())
        .expect("ecma:reflect.__proxyRevoke must be registered");

    crate::perf::register_host_fn(
        vm,
        "ecma:reflect",
        "proxyRevocable",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let handler = args.get(1).cloned().unwrap_or(Value::Undefined);

            let mut proxy = Object::new();
            proxy
                .properties
                .insert("__vybe_proxy_target".into(), target.clone());
            proxy
                .properties
                .insert("__vybe_proxy_handler".into(), handler);
            if let Value::Object(target_obj) = &target {
                if let Some(proto) = target_obj
                    .lock()
                    .unwrap()
                    .properties
                    .get("__proto__")
                    .cloned()
                {
                    proxy.properties.insert("__proto__".into(), proto);
                }
            }
            let proxy_value = Value::Object(vybe_runtime::heap::alloc(proxy));

            let mut revoke = Object::new();
            revoke.kind = ObjectKind::HostFunction(proxy_revoke_idx);
            revoke.properties.insert(
                "__bound_args".into(),
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                    proxy_value.clone(),
                ]))),
            );
            let revoke_value = Value::Object(vybe_runtime::heap::alloc(revoke));

            let mut result = Object::new();
            result.properties.insert("proxy".into(), proxy_value);
            result.properties.insert("revoke".into(), revoke_value);
            Value::Object(vybe_runtime::heap::alloc(result))
        }),
    );

    // Reflect.apply(target, thisArg, argsList) → result
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            let undefined = Value::Undefined;
            let target = args.first().unwrap_or(&undefined);
            // §28.1.1 step 1: IsCallable(target) — plain objects throw.
            let callable = crate::function::is_callable(target);
            if !callable {
                return throw_type_error(ctx, "Reflect.apply target is not a function");
            }
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            let mut inline_args: [Value; 8] = std::array::from_fn(|_| Value::Undefined);
            let invoke_args = match collect_apply_args_once(ctx, args.get(2), &mut inline_args) {
                Ok(values) => values,
                Err(error) => {
                    ctx.throw_value(error);
                    return Value::Undefined;
                }
            };
            invoke_with_explicit_this(ctx, target, this_arg, invoke_args.as_slice())
        }),
    );

    // Reflect.construct(target, argsList, newTarget?) → object
    //
    // §28.1.2 routes through target.[[Construct]]; for proxy exotic
    // objects that is the construct trap (§10.5.13), so delegate to the
    // shared dispatch in ecma::proxy.
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "construct",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let args_list = args.get(1).cloned().unwrap_or_else(|| {
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
            });
            let new_target = args.get(2).cloned();
            crate::proxy::construct_dispatch_with_new_target(ctx, &target, &args_list, new_target)
        }),
    );

    // Reflect.get(target, key, receiver?) → value
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "get",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let receiver = args.get(2).cloned().unwrap_or_else(|| target.clone());
            with_property_key(args.get(1), |key| reflect_get(ctx, &target, key, receiver))
        }),
    );

    // Reflect.set(target, key, value, receiver?) → bool (always true here)
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "set",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.get(3).cloned().unwrap_or_else(|| target.clone());
            with_property_key(args.get(1), |key| {
                reflect_set(ctx, &target, key, val, receiver)
            })
        }),
    );

    // Replaces the retired VM-internal `REF_IS_FUNC` opcode — matches its
    // EXACT semantics: true only for a user-defined `ObjectKind::Function`,
    // NOT host functions. The sole caller (the for-of custom-iterator gate)
    // relies on this narrow test: a built-in Map/Set `[Symbol.iterator]` is a
    // HostFunction and must fall through to native iteration, not the custom
    // lazy-iterator path. (Deliberately narrower than ECMA IsCallable §7.2.3.)
    crate::perf::register_host_fn(
        vm,
        "ecma:reflect",
        "isCallable",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let is_fn = matches!(
                args.first(),
                Some(Value::Object(o)) if matches!(
                    o.lock().unwrap().kind,
                    ObjectKind::Function(_)
                )
            );
            Value::I32(i32::from(is_fn))
        }),
    );

    // Reflect.has(target, key) → bool. Mirrors `key in target` (own + proto).
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "has",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            if let Some(Value::Object(obj)) = args.first() {
                return with_property_key(args.get(1), |key| {
                    // §28.1.8 on a proxy routes through the has trap (§10.5.7).
                    if let Some(proxy) = crate::proxy::is_proxy(args.first().unwrap()) {
                        if crate::proxy::proxy_is_revoked(&proxy) {
                            return throw_type_error(
                                ctx,
                                "Cannot perform 'has' on a revoked proxy",
                            );
                        }
                    }
                    if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
                        if let Some(trap) = proxy_trap(&handler, "has") {
                            let result = invoke_with_explicit_this(
                                ctx,
                                &trap,
                                handler,
                                &[proxy_target, crate::keys::string_value(key)],
                            );
                            return Value::Bool(result.as_bool());
                        }
                        if let Value::Object(t) = proxy_target {
                            return Value::Bool(proto_walk_has(&t, key));
                        }
                    }
                    // §10.4.2/§10.4.5: array & typed-array element indices are
                    // own properties.
                    {
                        let o = obj.lock().unwrap();
                        if let Some(i) = crate::keys::canonical_array_index_key(key) {
                            let i = i as usize;
                            match &o.kind {
                                ObjectKind::Array(elems) if i < elems.len() => {
                                    return Value::Bool(true);
                                }
                                ObjectKind::TypedArray(ta) if i < ta.length => {
                                    return Value::Bool(true);
                                }
                                _ => {}
                            }
                        }
                    }
                    Value::Bool(proto_walk_has(obj, key))
                });
            }
            Value::Bool(false)
        }),
    );

    // Reflect.deleteProperty(target, key) → bool
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "deleteProperty",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            if let Some(Value::Object(obj)) = args.first() {
                return with_property_key(args.get(1), |key| {
                    // §28.1.4 on a proxy: deleteProperty trap, else the target.
                    if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
                        if let Some(trap) = proxy_trap(&handler, "deleteProperty") {
                            let result = invoke_with_explicit_this(
                                ctx,
                                &trap,
                                handler,
                                &[proxy_target, crate::keys::string_value(key)],
                            );
                            return Value::Bool(result.as_bool());
                        }
                        if let Value::Object(t) = &proxy_target {
                            let mut o = t.lock().unwrap();
                            if o.properties.get("__vybe_frozen").is_some()
                                || o.properties.get("__vybe_sealed").is_some()
                                || is_nonconfig(&o, key)
                            {
                                return Value::Bool(false);
                            }
                            o.properties.shift_remove(key);
                            return Value::Bool(true);
                        }
                    }
                    let mut o = obj.lock().unwrap();
                    // §7.3.8: sealed/frozen objects have non-configurable
                    // properties — delete is refused.
                    if o.properties.get("__vybe_frozen").is_some()
                        || o.properties.get("__vybe_sealed").is_some()
                        || is_nonconfig(&o, key)
                    {
                        return Value::Bool(false);
                    }
                    o.properties.shift_remove(key);
                    Value::Bool(true)
                });
            }
            Value::Bool(true)
        }),
    );

    // Reflect.ownKeys(target) → Array of string keys.
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "ownKeys",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            // §28.1.10 routes through [[OwnPropertyKeys]] — proxies get
            // their ownKeys trap (or the target's keys when trapless).
            if let Some(v) = args.first() {
                if let Some(result) = crate::proxy::own_keys_dispatch(ctx, v) {
                    return result;
                }
            }
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let base_len = match &o.kind {
                    ObjectKind::Array(elems) => elems.len().saturating_add(1),
                    ObjectKind::TypedArray(ta) => ta.length,
                    _ => 0,
                };
                let mut keys: Vec<Value> = Vec::with_capacity(base_len + o.properties.len());
                // §10.4.2.4 OwnPropertyKeys: integer indices first, then
                // "length", then other string keys. Elements live in the
                // kind, not in `properties`.
                match &o.kind {
                    ObjectKind::Array(elems) => {
                        for i in 0..elems.len() {
                            crate::keys::with_index_key(i, |key| {
                                keys.push(crate::keys::string_value(key));
                            });
                        }
                        keys.push(crate::keys::string_value("length"));
                    }
                    ObjectKind::TypedArray(ta) => {
                        for i in 0..ta.length {
                            crate::keys::with_index_key(i, |key| {
                                keys.push(crate::keys::string_value(key));
                            });
                        }
                    }
                    _ => {}
                }
                for key in ordered_own_string_keys(&o) {
                    if key != "length" || !matches!(o.kind, ObjectKind::Array(_)) {
                        keys.push(crate::keys::string_value(key.as_str()));
                    }
                }
                if let Some(Value::Object(sym_keys)) = o.properties.get("__sym_keys") {
                    if let ObjectKind::Array(sym_entries) = &sym_keys.lock().unwrap().kind {
                        keys.extend(sym_entries.iter().cloned());
                    }
                }
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys)));
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![])))
        }),
    );

    // Reflect.getOwnPropertyDescriptor(target, key) → descriptor object | undefined
    crate::perf::register_host_fn(
        vm,
        "ecma:reflect",
        "getOwnPropertyDescriptor",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                return with_property_key(args.get(1), |key| {
                    let o = obj.lock().unwrap();
                    if let Some(val) = o.properties.get(key) {
                        let mut desc = Object::new();
                        desc.properties.insert("value".into(), val.clone());
                        desc.properties.insert("writable".into(), Value::Bool(true));
                        desc.properties
                            .insert("enumerable".into(), Value::Bool(!key.starts_with("__")));
                        desc.properties
                            .insert("configurable".into(), Value::Bool(true));
                        return Value::Object(vybe_runtime::heap::alloc(desc));
                    }
                    Value::Undefined
                });
            }
            Value::Undefined
        }),
    );

    // Reflect.defineProperty(target, key, attrs) → bool
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "defineProperty",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            // §28.1.3 step 1: target must be an Object.
            if !matches!(args.first(), Some(Value::Object(_))) {
                return throw_type_error(ctx, "Reflect.defineProperty called on non-object");
            }
            let (val, enumerable, writable, configurable) =
                if let Some(Value::Object(attrs)) = args.get(2) {
                    let a = attrs.lock().unwrap();
                    (
                        a.properties.get("value").cloned(),
                        a.properties
                            .get("enumerable")
                            .map(|v| v.as_bool())
                            .unwrap_or(false),
                        a.properties.get("writable").map(|v| v.as_bool()),
                        a.properties
                            .get("configurable")
                            .map(|v| v.as_bool())
                            .unwrap_or(false),
                    )
                } else {
                    (None, false, None, false)
                };
            if let Some(Value::Object(obj)) = args.first() {
                return with_property_key(args.get(1), |key| {
                    track_key(obj, key);
                    if !enumerable {
                        track_nonenum(obj, key);
                    }
                    let mut o = obj.lock().unwrap();
                    if o.properties.contains_key(key) && is_nonconfig(&o, key) {
                        return Value::Bool(false);
                    }
                    if is_not_extensible(&o) && !o.properties.contains_key(key) {
                        return Value::Bool(false);
                    }
                    if let Some(v) = val {
                        o.properties.insert(key.to_string(), v);
                        if matches!(writable, Some(false) | None) {
                            install_noop_setter(&mut o, key);
                        }
                    } else if !o.properties.contains_key(key) {
                        o.properties.insert(key.to_string(), Value::Undefined);
                    }
                    drop(o);
                    if !configurable {
                        track_nonconfig(obj, key);
                    }
                    Value::Bool(true)
                });
            }
            Value::Bool(false)
        }),
    );

    // Reflect.getPrototypeOf(target) → Object | null
    //
    // Vybe stores the prototype under `__proto__`; missing → null.
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "getPrototypeOf",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            // §28.1.7 step 1: unlike Object.getPrototypeOf, a non-object
            // target is a TypeError here — Reflect does not coerce.
            let Some(value @ Value::Object(_)) = args.first() else {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Reflect.getPrototypeOf called on non-object",
                ));
                return Value::Undefined;
            };
            crate::object::get_prototype_of(ctx, value).unwrap_or(Value::Undefined)
        }),
    );

    // Reflect.setPrototypeOf(target, proto) → bool
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "setPrototypeOf",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            // §28.1.14 steps 1-2: non-object target, or a prototype that is
            // neither object nor null, is a TypeError.
            let Some(value @ Value::Object(_)) = args.first() else {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Reflect.setPrototypeOf called on non-object",
                ));
                return Value::Undefined;
            };
            let proto = args.get(1).cloned().unwrap_or(Value::Null);
            if !matches!(proto, Value::Object(_) | Value::Null) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Object prototype may only be an Object or null",
                ));
                return Value::Undefined;
            }
            // Step 3 hands the SUCCESS FLAG straight back — where
            // Object.setPrototypeOf would throw, this returns false.
            match crate::object::set_prototype_of(ctx, value, &proto) {
                None => Value::Undefined,
                Some(success) => Value::Bool(success),
            }
        }),
    );

    // Reflect.isExtensible(target) → bool.
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "isExtensible",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            let Some(value) = args.first() else {
                return throw_type_error(ctx, "Reflect.isExtensible called on non-object");
            };
            if let Some(proxy) = crate::proxy::is_proxy(value) {
                if crate::proxy::proxy_is_revoked(&proxy) {
                    return throw_type_error(
                        ctx,
                        "Cannot perform 'isExtensible' on a proxy that has been revoked",
                    );
                }
                let Some((target, handler)) = proxy_target_and_handler(&proxy) else {
                    return Value::Bool(false);
                };
                let target_extensible = crate::object::value_is_extensible(&target);
                let reported = if let Some(trap) = proxy_trap(&handler, "isExtensible") {
                    crate::boolean::to_boolean(&invoke_with_explicit_this(
                        ctx,
                        &trap,
                        handler,
                        &[target.clone()],
                    ))
                } else {
                    target_extensible
                };
                if reported != target_extensible {
                    return throw_type_error(
                        ctx,
                        "Proxy isExtensible trap result does not match target",
                    );
                }
                return Value::Bool(reported);
            }
            if let Value::Object(obj) = value {
                let o = obj.lock().unwrap();
                // §7.3.15 SetIntegrityLevel: seal/freeze also call
                // [[PreventExtensions]] — a sealed object is not extensible.
                let sealed = o.properties.get("__vybe_sealed").is_some()
                    || o.properties.get("__vybe_frozen").is_some();
                return Value::Bool(!sealed && !is_not_extensible(&o));
            }
            throw_type_error(ctx, "Reflect.isExtensible called on non-object")
        }),
    );

    // Reflect.preventExtensions(target) → bool
    crate::perf::register_free_fn(
        vm,
        "ecma:reflect",
        "preventExtensions",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // ⛔ Every `ecma:reflect` entry is reached as a METHOD
            // (`Reflect.set(...)`), so under `ReceiverAbi::Parameter`
            // argument 0 is the receiver. Shadowing `args` once here
            // keeps every positional read below correct under both
            // bindings — `user_args` skips nothing when there is no
            // receiver slot. Reading at fixed indices made
            // `Reflect.set(t,"v",5)` pass the Reflect object as the
            // target and the target as the value.
            let args = ctx.user_args(args, 0);
            let Some(value) = args.first() else {
                return throw_type_error(ctx, "Reflect.preventExtensions called on non-object");
            };
            if let Some(proxy) = crate::proxy::is_proxy(value) {
                if crate::proxy::proxy_is_revoked(&proxy) {
                    return throw_type_error(
                        ctx,
                        "Cannot perform 'preventExtensions' on a proxy that has been revoked",
                    );
                }
                let Some((target, handler)) = proxy_target_and_handler(&proxy) else {
                    return Value::Bool(false);
                };
                let success = if let Some(trap) = proxy_trap(&handler, "preventExtensions") {
                    crate::boolean::to_boolean(&invoke_with_explicit_this(
                        ctx,
                        &trap,
                        handler,
                        &[target.clone()],
                    ))
                } else {
                    if let Value::Object(target_obj) = &target {
                        mark_not_extensible(&mut target_obj.lock().unwrap());
                    }
                    true
                };
                if success && crate::object::value_is_extensible(&target) {
                    return throw_type_error(
                        ctx,
                        "Proxy preventExtensions trap returned true but target is still extensible",
                    );
                }
                return Value::Bool(success);
            }
            if let Value::Object(obj) = value {
                mark_not_extensible(&mut obj.lock().unwrap());
                return Value::Bool(true);
            }
            throw_type_error(ctx, "Reflect.preventExtensions called on non-object")
        }),
    );
}
