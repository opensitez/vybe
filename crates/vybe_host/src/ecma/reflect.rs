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

use crate::ecma::function::invoke_with_explicit_this;
use crate::ecma::object::{
    install_noop_setter, is_nonconfig, is_not_extensible, js_prototype_of, mark_not_extensible,
    ordered_own_string_keys, proto_walk_get, proxy_target_and_handler, proxy_trap, track_key,
    track_nonconfig, track_nonenum,
};
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};


// §28.1 step 1 of most Reflect ops: "If target is not an Object, throw a
// TypeError". Thrown as a real error object so `e instanceof TypeError`
// holds in the catcher.
fn throw_type_error(ctx: &mut HostContext, message: &str) -> Value {
    ctx.throw_value(crate::ecma::error::new_error("TypeError", message));
    Value::Undefined
}

fn numeric_index(key: &str) -> Option<usize> {
    key.parse::<usize>().ok()
}

/// §28.1.5 Reflect.get — proxy trap, typed-array elements, and a
/// [[Get]]-shaped prototype walk that invokes accessors with `receiver`.
fn reflect_get(ctx: &mut HostContext, target: &Value, key: &str, receiver: Value) -> Value {
    let Value::Object(obj) = target else {
        return throw_type_error(ctx, "Reflect.get called on non-object");
    };
    if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
        if let Some(trap) = proxy_trap(&handler, "get") {
            return invoke_with_explicit_this(
                ctx,
                &trap,
                handler,
                &[
                    proxy_target,
                    Value::String(Arc::from(key)),
                    receiver,
                ],
            );
        }
        return reflect_get(ctx, &proxy_target, key, receiver);
    }
    {
        let o = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref ta) = o.kind {
            if let Some(i) = numeric_index(key) {
                if i < ta.length {
                    return crate::ecma::typedarray::read_element(ta, i);
                }
                return Value::Undefined;
            }
        }
    }
    // [[Get]] walk: accessor (getter) beats data at each level; the
    // getter runs with `this = receiver` (§10.1.8.1 step 8).
    let getter_key = format!("__get_{}", key);
    let mut current = Some(obj.clone());
    while let Some(node) = current {
        let (getter, data, next) = {
            let o = node.lock().unwrap();
            (
                o.properties.get(&getter_key).cloned(),
                o.properties.get(key).cloned(),
                o.properties.get("__proto__").cloned(),
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
                return invoke_with_explicit_this(ctx, &g, receiver, &[]);
            }
            return invoke_with_explicit_this(ctx, &g, receiver.clone(), &[receiver]);
        }
        if let Some(v) = data {
            return v;
        }
        current = match next {
            Some(Value::Object(p)) => Some(p),
            _ => None,
        };
    }
    Value::Undefined
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
                &[
                    proxy_target,
                    Value::String(Arc::from(key)),
                    val,
                    receiver,
                ],
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
                    crate::ecma::typedarray::write_element(ta, i, &val);
                    return Value::Bool(true);
                }
                return Value::Bool(false);
            }
        }
    }
    // Walk for an accessor: a REAL setter (compiled Function) runs with
    // `this = receiver` (§10.1.9.2); the noop setter installed for
    // non-writable data properties (a HostFunction) means reject. A
    // getter with no setter at the same level also rejects.
    let setter_key = format!("__set_{}", key);
    let getter_key = format!("__get_{}", key);
    let mut current = Some(obj.clone());
    while let Some(node) = current {
        let (setter, has_getter, has_data, next) = {
            let o = node.lock().unwrap();
            (
                o.properties.get(&setter_key).cloned(),
                o.properties.contains_key(&getter_key),
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
                if arity >= 2 {
                    invoke_with_explicit_this(
                        ctx,
                        &st,
                        receiver.clone(),
                        &[receiver, val],
                    );
                } else {
                    invoke_with_explicit_this(ctx, &st, receiver, &[val]);
                }
                return Value::Bool(true);
            }
            // noop setter (HostFunction) = non-writable data property
            return Value::Bool(false);
        }
        if has_getter {
            // §10.1.9.2 step 6.c: accessor without a [[Set]] → false
            return Value::Bool(false);
        }
        if has_data {
            break; // data property found — assign below
        }
        current = match next {
            Some(Value::Object(p)) => Some(p),
            _ => None,
        };
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
    vm.register_host_fn(
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
        .get(&("ecma:reflect".to_string(), "__proxyRevoke".to_string()))
        .expect("ecma:reflect.__proxyRevoke must be registered");

    vm.register_host_fn(
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
            let proxy_value = Value::Object(Arc::new(Mutex::new(proxy)));

            let mut revoke = Object::new();
            revoke.kind = ObjectKind::HostFunction(proxy_revoke_idx);
            revoke.properties.insert(
                "__bound_args".into(),
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                    proxy_value.clone(),
                ])))),
            );
            let revoke_value = Value::Object(Arc::new(Mutex::new(revoke)));

            let mut result = Object::new();
            result.properties.insert("proxy".into(), proxy_value);
            result.properties.insert("revoke".into(), revoke_value);
            Value::Object(Arc::new(Mutex::new(result)))
        }),
    );

    // Reflect.apply(target, thisArg, argsList) → result
    vm.register_host_fn(
        "ecma:reflect",
        "apply",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            // §28.1.1 step 1: IsCallable(target) — plain objects throw.
            let callable = matches!(
                &target,
                Value::Object(o) if matches!(
                    o.lock().unwrap().kind,
                    ObjectKind::Function(_) | ObjectKind::HostFunction(_)
                )
            );
            if !callable {
                return throw_type_error(ctx, "Reflect.apply target is not a function");
            }
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            let mut invoke_args: Vec<Value> = Vec::new();
            if let Some(Value::Object(arr)) = args.get(2) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    invoke_args.extend(v.iter().cloned());
                }
                // Array-like {0:…, 1:…, length: n} per CreateListFromArrayLike.
                if invoke_args.is_empty() {
                    if let Some(len) = o.properties.get("length").map(|v| v.as_i32()) {
                        for i in 0..len.max(0) {
                            invoke_args.push(
                                o.properties
                                    .get(&i.to_string())
                                    .cloned()
                                    .unwrap_or(Value::Undefined),
                            );
                        }
                    }
                }
            }
            let result = invoke_with_explicit_this(ctx, &target, this_arg, &invoke_args);
            if matches!(result, Value::Null) {
                Value::Undefined
            } else {
                result
            }
        }),
    );

    // Reflect.construct(target, argsList, newTarget?) → object
    //
    // §28.1.2 routes through target.[[Construct]]; for proxy exotic
    // objects that is the construct trap (§10.5.13), so delegate to the
    // shared dispatch in ecma::proxy.
    vm.register_host_fn(
        "ecma:reflect",
        "construct",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let args_list = args.get(1).cloned().unwrap_or_else(|| {
                Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
            });
            let new_target = args.get(2).cloned();
            crate::ecma::proxy::construct_dispatch_with_new_target(
                ctx,
                &target,
                &args_list,
                new_target,
            )
        }),
    );

    // Reflect.get(target, key, receiver?) → value
    vm.register_host_fn(
        "ecma:reflect",
        "get",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let receiver = args.get(2).cloned().unwrap_or_else(|| target.clone());
            reflect_get(ctx, &target, &key, receiver)
        }),
    );

    // Reflect.set(target, key, value, receiver?) → bool (always true here)
    vm.register_host_fn(
        "ecma:reflect",
        "set",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let target = args.first().cloned().unwrap_or(Value::Undefined);
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            let receiver = args.get(3).cloned().unwrap_or_else(|| target.clone());
            reflect_set(ctx, &target, &key, val, receiver)
        }),
    );

    // Reflect.has(target, key) → bool. Mirrors `key in target` (own + proto).
    vm.register_host_fn(
        "ecma:reflect",
        "has",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Some(Value::Object(obj)) = args.first() {
                // §28.1.8 on a proxy routes through the has trap (§10.5.7).
                if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
                    if let Some(trap) = proxy_trap(&handler, "has") {
                        let result = invoke_with_explicit_this(
                            ctx,
                            &trap,
                            handler,
                            &[proxy_target, Value::String(Arc::from(key.as_str()))],
                        );
                        return Value::Bool(result.as_bool());
                    }
                    if let Value::Object(t) = proxy_target {
                        return Value::Bool(proto_walk_get(&t, &key).is_some());
                    }
                }
                // §10.4.2/§10.4.5: array & typed-array element indices are
                // own properties.
                {
                    let o = obj.lock().unwrap();
                    if let Some(i) = key.parse::<usize>().ok() {
                        match &o.kind {
                            ObjectKind::Array(elems) if i < elems.len() => {
                                return Value::Bool(true)
                            }
                            ObjectKind::TypedArray(ta) if i < ta.length => {
                                return Value::Bool(true)
                            }
                            _ => {}
                        }
                    }
                }
                return Value::Bool(proto_walk_get(obj, &key).is_some());
            }
            Value::Bool(false)
        }),
    );

    // Reflect.deleteProperty(target, key) → bool
    vm.register_host_fn(
        "ecma:reflect",
        "deleteProperty",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Some(Value::Object(obj)) = args.first() {
                // §28.1.4 on a proxy: deleteProperty trap, else the target.
                if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
                    if let Some(trap) = proxy_trap(&handler, "deleteProperty") {
                        let result = invoke_with_explicit_this(
                            ctx,
                            &trap,
                            handler,
                            &[proxy_target, Value::String(Arc::from(key.as_str()))],
                        );
                        return Value::Bool(result.as_bool());
                    }
                    if let Value::Object(t) = &proxy_target {
                        let mut o = t.lock().unwrap();
                        if o.properties.get("__vybe_frozen").is_some()
                            || o.properties.get("__vybe_sealed").is_some()
                            || is_nonconfig(&o, &key)
                        {
                            return Value::Bool(false);
                        }
                        o.properties.remove(&key);
                        return Value::Bool(true);
                    }
                }
                let mut o = obj.lock().unwrap();
                // §7.3.8: sealed/frozen objects have non-configurable
                // properties — delete is refused.
                if o.properties.get("__vybe_frozen").is_some()
                    || o.properties.get("__vybe_sealed").is_some()
                    || is_nonconfig(&o, &key)
                {
                    return Value::Bool(false);
                }
                o.properties.remove(&key);
                return Value::Bool(true);
            }
            Value::Bool(true)
        }),
    );

    // Reflect.ownKeys(target) → Array of string keys.
    vm.register_host_fn(
        "ecma:reflect",
        "ownKeys",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // §28.1.10 routes through [[OwnPropertyKeys]] — proxies get
            // their ownKeys trap (or the target's keys when trapless).
            if let Some(v) = args.first() {
                if let Some(result) = crate::ecma::proxy::own_keys_dispatch(ctx, v) {
                    return result;
                }
            }
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let mut keys: Vec<Value> = Vec::new();
                // §10.4.2.4 OwnPropertyKeys: integer indices first, then
                // "length", then other string keys. Elements live in the
                // kind, not in `properties`.
                match &o.kind {
                    ObjectKind::Array(elems) => {
                        for i in 0..elems.len() {
                            keys.push(Value::String(Arc::from(i.to_string().as_str())));
                        }
                        keys.push(Value::String(Arc::from("length")));
                    }
                    ObjectKind::TypedArray(ta) => {
                        for i in 0..ta.length {
                            keys.push(Value::String(Arc::from(i.to_string().as_str())));
                        }
                    }
                    _ => {}
                }
                keys.extend(
                    ordered_own_string_keys(&o)
                        .into_iter()
                        .filter(|k| k != "length" || !matches!(o.kind, ObjectKind::Array(_)))
                        .map(|k| Value::String(Arc::from(k.as_str()))),
                );
                if let Some(Value::Object(sym_keys)) = o.properties.get("__sym_keys") {
                    if let ObjectKind::Array(sym_entries) = &sym_keys.lock().unwrap().kind {
                        keys.extend(sym_entries.iter().cloned());
                    }
                }
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }),
    );

    // Reflect.getOwnPropertyDescriptor(target, key) → descriptor object | undefined
    vm.register_host_fn(
        "ecma:reflect",
        "getOwnPropertyDescriptor",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                if let Some(val) = o.properties.get(&key) {
                    let mut desc = Object::new();
                    desc.properties.insert("value".into(), val.clone());
                    desc.properties.insert("writable".into(), Value::Bool(true));
                    desc.properties
                        .insert("enumerable".into(), Value::Bool(!key.starts_with("__")));
                    desc.properties
                        .insert("configurable".into(), Value::Bool(true));
                    return Value::Object(Arc::new(Mutex::new(desc)));
                }
            }
            Value::Undefined
        }),
    );

    // Reflect.defineProperty(target, key, attrs) → bool
    vm.register_host_fn(
        "ecma:reflect",
        "defineProperty",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // §28.1.3 step 1: target must be an Object.
            if !matches!(args.first(), Some(Value::Object(_))) {
                return throw_type_error(ctx, "Reflect.defineProperty called on non-object");
            }
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
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
                track_key(obj, &key);
                if !enumerable {
                    track_nonenum(obj, &key);
                }
                let mut o = obj.lock().unwrap();
                if o.properties.contains_key(&key) && is_nonconfig(&o, &key) {
                    return Value::Bool(false);
                }
                if is_not_extensible(&o) && !o.properties.contains_key(&key) {
                    return Value::Bool(false);
                }
                if let Some(v) = val {
                    o.properties.insert(key.clone(), v);
                    if matches!(writable, Some(false) | None) {
                        install_noop_setter(&mut o, &key);
                    }
                } else if !o.properties.contains_key(&key) {
                    o.properties.insert(key.clone(), Value::Undefined);
                }
                drop(o);
                if !configurable {
                    track_nonconfig(obj, &key);
                }
                return Value::Bool(true);
            }
            Value::Bool(false)
        }),
    );

    // Reflect.getPrototypeOf(target) → Object | null
    //
    // Vybe stores the prototype under `__proto__`; missing → null.
    vm.register_host_fn(
        "ecma:reflect",
        "getPrototypeOf",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                if let Some((proxy_target, handler)) = proxy_target_and_handler(obj) {
                    if let Some(trap) = proxy_trap(&handler, "getPrototypeOf") {
                        return invoke_with_explicit_this(ctx, &trap, handler, &[proxy_target]);
                    }
                    return js_prototype_of(&proxy_target);
                }
                return js_prototype_of(&Value::Object(obj.clone()));
            }
            Value::Null
        }),
    );

    // Reflect.setPrototypeOf(target, proto) → bool
    vm.register_host_fn(
        "ecma:reflect",
        "setPrototypeOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let proto = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                o.properties.insert("__proto__".into(), proto);
                return Value::Bool(true);
            }
            Value::Bool(false)
        }),
    );

    // Reflect.isExtensible(target) → bool.
    vm.register_host_fn(
        "ecma:reflect",
        "isExtensible",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                // §7.3.15 SetIntegrityLevel: seal/freeze also call
                // [[PreventExtensions]] — a sealed object is not extensible.
                let sealed = o.properties.get("__vybe_sealed").is_some()
                    || o.properties.get("__vybe_frozen").is_some();
                return Value::Bool(!sealed && !is_not_extensible(&o));
            }
            Value::Bool(false)
        }),
    );

    // Reflect.preventExtensions(target) → bool
    vm.register_host_fn(
        "ecma:reflect",
        "preventExtensions",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                mark_not_extensible(&mut o);
                return Value::Bool(true);
            }
            Value::Bool(false)
        }),
    );
}
