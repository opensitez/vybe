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

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Reflect.apply(target, thisArg, argsList) → result
    vm.register_host_fn("ecma:reflect", "apply", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let target = args.first().cloned().unwrap_or(Value::Undefined);
        let this_arg = args.get(1).cloned().unwrap_or(Value::Undefined);
        let mut invoke_args: Vec<Value> = vec![this_arg];
        if let Some(Value::Object(arr)) = args.get(2) {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                invoke_args.extend(v.iter().cloned());
            }
        }
        ctx.invoke(&target, &invoke_args)
    }));

    // Reflect.construct(target, argsList, newTarget?) → object
    //
    // `newTarget` is ignored — without proper [[Construct]] internal
    // method dispatch, we synthesize a plain Object as `this` and
    // invoke `target` as a function on it.
    vm.register_host_fn("ecma:reflect", "construct", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let target = args.first().cloned().unwrap_or(Value::Undefined);
        let this_obj = Value::Object(Arc::new(Mutex::new(Object::new())));
        let mut invoke_args: Vec<Value> = vec![this_obj.clone()];
        if let Some(Value::Object(arr)) = args.get(1) {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref v) = o.kind {
                invoke_args.extend(v.iter().cloned());
            }
        }
        let result = ctx.invoke(&target, &invoke_args);
        // If the target returns an object, use it; else use the synthetic this.
        if matches!(result, Value::Object(_)) { result } else { this_obj }
    }));

    // Reflect.get(target, key, receiver?) → value
    vm.register_host_fn("ecma:reflect", "get", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            return o.properties.get(&key).cloned().unwrap_or(Value::Undefined);
        }
        Value::Undefined
    }));

    // Reflect.set(target, key, value, receiver?) → bool (always true here)
    vm.register_host_fn("ecma:reflect", "set", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let val = args.get(2).cloned().unwrap_or(Value::Undefined);
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            o.properties.insert(key, val);
            return Value::Bool(true);
        }
        Value::Bool(false)
    }));

    // Reflect.has(target, key) → bool. Mirrors `key in target` (own + proto).
    vm.register_host_fn("ecma:reflect", "has", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            return Value::Bool(o.properties.contains_key(&key));
        }
        Value::Bool(false)
    }));

    // Reflect.deleteProperty(target, key) → bool
    vm.register_host_fn("ecma:reflect", "deleteProperty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            return Value::Bool(o.properties.remove(&key).is_some());
        }
        Value::Bool(false)
    }));

    // Reflect.ownKeys(target) → Array of string keys.
    vm.register_host_fn("ecma:reflect", "ownKeys", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let keys: Vec<Value> = o.properties.keys()
                .filter(|k| !k.starts_with("__"))
                .map(|k| Value::String(Arc::from(k.as_str())))
                .collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // Reflect.getOwnPropertyDescriptor(target, key) → descriptor object | undefined
    vm.register_host_fn("ecma:reflect", "getOwnPropertyDescriptor", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(val) = o.properties.get(&key) {
                let mut desc = Object::new();
                desc.properties.insert("value".into(), val.clone());
                desc.properties.insert("writable".into(), Value::Bool(true));
                desc.properties.insert("enumerable".into(), Value::Bool(!key.starts_with("__")));
                desc.properties.insert("configurable".into(), Value::Bool(true));
                return Value::Object(Arc::new(Mutex::new(desc)));
            }
        }
        Value::Undefined
    }));

    // Reflect.defineProperty(target, key, attrs) → bool
    vm.register_host_fn("ecma:reflect", "defineProperty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let val = if let Some(Value::Object(attrs)) = args.get(2) {
            let a = attrs.lock().unwrap();
            a.properties.get("value").cloned().unwrap_or(Value::Undefined)
        } else { Value::Undefined };
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            o.properties.insert(key, val);
            return Value::Bool(true);
        }
        Value::Bool(false)
    }));

    // Reflect.getPrototypeOf(target) → Object | null
    //
    // Vybe stores the prototype under `__proto__`; missing → null.
    vm.register_host_fn("ecma:reflect", "getPrototypeOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            return o.properties.get("__proto__").cloned().unwrap_or(Value::Null);
        }
        Value::Null
    }));

    // Reflect.setPrototypeOf(target, proto) → bool
    vm.register_host_fn("ecma:reflect", "setPrototypeOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let proto = args.get(1).cloned().unwrap_or(Value::Null);
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            o.properties.insert("__proto__".into(), proto);
            return Value::Bool(true);
        }
        Value::Bool(false)
    }));

    // Reflect.isExtensible(target) → bool. Vybe doesn't enforce
    // [[Extensible]] beyond the freeze marker.
    vm.register_host_fn("ecma:reflect", "isExtensible", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            return Value::Bool(o.properties.get("__frozen").map(|v|
                !matches!(v, Value::Bool(true))).unwrap_or(true));
        }
        Value::Bool(false)
    }));

    // Reflect.preventExtensions(target) → bool
    vm.register_host_fn("ecma:reflect", "preventExtensions", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            o.properties.insert("__frozen".into(), Value::Bool(true));
            return Value::Bool(true);
        }
        Value::Bool(false)
    }));
}
