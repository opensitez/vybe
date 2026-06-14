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
    install_noop_setter, is_nonconfig, is_not_extensible, mark_not_extensible,
    ordered_own_string_keys, proto_walk_get, track_key, track_nonconfig, track_nonenum,
};
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

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
            let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
            let mut invoke_args: Vec<Value> = Vec::new();
            if let Some(Value::Object(arr)) = args.get(2) {
                let o = arr.lock().unwrap();
                if let ObjectKind::Array(ref v) = o.kind {
                    invoke_args.extend(v.iter().cloned());
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
            crate::ecma::proxy::construct_dispatch(ctx, &target, &args_list)
        }),
    );

    // Reflect.get(target, key, receiver?) → value
    vm.register_host_fn(
        "ecma:reflect",
        "get",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Some(Value::Object(obj)) = args.first() {
                return proto_walk_get(obj, &key).unwrap_or(Value::Undefined);
            }
            Value::Undefined
        }),
    );

    // Reflect.set(target, key, value, receiver?) → bool (always true here)
    vm.register_host_fn(
        "ecma:reflect",
        "set",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let val = args.get(2).cloned().unwrap_or(Value::Undefined);
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                if o.properties.contains_key(&format!("__set_{}", key)) {
                    return Value::Bool(false);
                }
                if is_not_extensible(&o) && !o.properties.contains_key(&key) {
                    return Value::Bool(false);
                }
                o.properties.insert(key, val);
                return Value::Bool(true);
            }
            Value::Bool(false)
        }),
    );

    // Reflect.has(target, key) → bool. Mirrors `key in target` (own + proto).
    vm.register_host_fn(
        "ecma:reflect",
        "has",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Some(Value::Object(obj)) = args.first() {
                return Value::Bool(proto_walk_get(obj, &key).is_some());
            }
            Value::Bool(false)
        }),
    );

    // Reflect.deleteProperty(target, key) → bool
    vm.register_host_fn(
        "ecma:reflect",
        "deleteProperty",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            if let Some(Value::Object(obj)) = args.first() {
                let mut o = obj.lock().unwrap();
                if is_nonconfig(&o, &key) {
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let mut keys: Vec<Value> = ordered_own_string_keys(&o)
                    .into_iter()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                return o
                    .properties
                    .get("__proto__")
                    .cloned()
                    .unwrap_or(Value::Null);
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
                return Value::Bool(!is_not_extensible(&o));
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
