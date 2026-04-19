//! # `wasm:js-object` host handlers
//!
//! Native Rust impls of `Object.*` statics and `Object.prototype.*` per
//! ECMA-262 §20.1, satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_object_builtins.rs`.
//!
//! Storage: our existing `Object { properties: HashMap<String, Value>,
//! kind: ObjectKind::Ordinary, … }`. Prototype chain is walked via a
//! magic `__proto__` property — same convention the VB / JS compilers
//! already use for class inheritance.
//!
//! This file also serves PHP's `array` (ordered string-or-int-keyed
//! dictionary) through the Vybe-specific `appendAutoKey` extension.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

/// Magic property name used to mark an object as frozen / sealed /
/// non-extensible. Matches existing vybe:object module's convention.
const FROZEN_MARK: &str = "__vybe_frozen";
const SEALED_MARK: &str = "__vybe_sealed";
const EXTENSIBLE_MARK: &str = "__vybe_extensible"; // absence means extensible
const PROTO_KEY: &str = "__proto__";
/// PHP-array next-int-key tracker. Used by `appendAutoKey`.
const NEXT_INT_KEY: &str = "__vybe_next_int_key";

fn is_object(v: &Value) -> bool {
    matches!(v, Value::Object(_))
}

fn obj_of(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        Some(obj.clone())
    } else {
        None
    }
}

fn key_string(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        _ => format!("{}", v),
    }
}

/// Walk the prototype chain looking for `key`. Returns the value if
/// found at any depth, `None` if not present in the whole chain.
fn proto_walk_get(obj: &Arc<Mutex<Object>>, key: &str) -> Option<Value> {
    let mut current = obj.clone();
    loop {
        let o = current.lock().unwrap();
        if let Some(v) = o.properties.get(key) {
            return Some(v.clone());
        }
        let proto = o.properties.get(PROTO_KEY).cloned();
        drop(o);
        match proto {
            Some(Value::Object(p)) => {
                if Arc::ptr_eq(&p, &current) {
                    // Cycle or self-proto; bail.
                    return None;
                }
                current = p;
            }
            _ => return None,
        }
    }
}

pub fn register(vm: &mut VM) {
    register_construction(vm);
    register_access(vm);
    register_enumeration(vm);
    register_descriptors(vm);
    register_prototype(vm);
    register_locking(vm);
    register_comparison(vm);
    register_prototype_methods(vm);
    register_php_extensions(vm);
}

// ── Construction ──────────────────────────────────────────────────────

fn register_construction(vm: &mut VM) {
    vm.register_host_fn("wasm:js-object", "new",
        Box::new(|_ctx, _args| {
            Value::Object(Arc::new(Mutex::new(Object::new())))
        }));

    // create(proto) -> new obj with prototype link
    vm.register_host_fn("wasm:js-object", "create",
        Box::new(|_ctx, args| {
            let mut obj = Object::new();
            if let Some(proto @ Value::Object(_)) = args.first() {
                obj.properties.insert(PROTO_KEY.into(), proto.clone());
            } else if matches!(args.first(), Some(Value::Null)) {
                // Object.create(null) — no prototype chain
                obj.properties.insert(PROTO_KEY.into(), Value::Null);
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // fromEntries(iterable) -> new obj
    vm.register_host_fn("wasm:js-object", "fromEntries",
        Box::new(|_ctx, args| {
            let mut obj = Object::new();
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref pairs) = s.kind {
                    for pair in pairs {
                        if let Value::Object(p) = pair {
                            let pl = p.lock().unwrap();
                            if let ObjectKind::Array(ref kv) = pl.kind {
                                if kv.len() >= 2 {
                                    obj.properties.insert(key_string(&kv[0]), kv[1].clone());
                                }
                            }
                        }
                    }
                }
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // assign(target, source) -> target (pairwise; multi-source chains)
    vm.register_host_fn("wasm:js-object", "assign",
        Box::new(|_ctx, args| {
            if let (Some(target), Some(source)) = (args.first(), args.get(1)) {
                if let (Value::Object(t), Value::Object(s)) = (target, source) {
                    let src = s.lock().unwrap();
                    let props: Vec<(String, Value)> = src.properties.iter()
                        .filter(|(k, _)| !k.starts_with("__"))
                        .map(|(k, v)| (k.clone(), v.clone()))
                        .collect();
                    drop(src);
                    let mut tgt = t.lock().unwrap();
                    for (k, v) in props {
                        tgt.properties.insert(k, v);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Null)
        }));
}

// ── Property access ───────────────────────────────────────────────────

fn register_access(vm: &mut VM) {
    // get(obj, key) -> value (walks prototype chain)
    vm.register_host_fn("wasm:js-object", "get",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                if let Some(v) = proto_walk_get(&obj, &key) {
                    return v;
                }
            }
            Value::Undefined
        }));

    // set(obj, key, value) -> ()
    vm.register_host_fn("wasm:js-object", "set",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut o = obj.lock().unwrap();
                // Per spec: writes fail silently on frozen / non-extensible
                if o.properties.get(FROZEN_MARK).is_some() {
                    return Value::Null;
                }
                o.properties.insert(key, val);
            }
            Value::Null
        }));

    // has(obj, key) -> i32 (walks prototype chain, returns 1/0)
    vm.register_host_fn("wasm:js-object", "has",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                return Value::I32(if proto_walk_get(&obj, &key).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    // hasOwn(obj, key) -> i32 (own-only, no prototype walk)
    vm.register_host_fn("wasm:js-object", "hasOwn",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                return Value::I32(if o.properties.contains_key(&key) { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    // delete(obj, key) -> i32 (1 if deleted)
    vm.register_host_fn("wasm:js-object", "delete",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let mut o = obj.lock().unwrap();
                if o.properties.get(SEALED_MARK).is_some() {
                    return Value::I32(0);
                }
                return Value::I32(if o.properties.remove(&key).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));
}

// ── Enumeration ───────────────────────────────────────────────────────

fn register_enumeration(vm: &mut VM) {
    fn own_keys(obj: &Object) -> Vec<String> {
        obj.properties
            .keys()
            .filter(|k| !k.starts_with("__"))
            .cloned()
            .collect()
    }

    vm.register_host_fn("wasm:js-object", "keys",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let keys: Vec<Value> = own_keys(&o).into_iter()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-object", "values",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let values: Vec<Value> = own_keys(&o).into_iter()
                    .filter_map(|k| o.properties.get(&k).cloned())
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(values))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-object", "entries",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let entries: Vec<Value> = own_keys(&o).into_iter()
                    .filter_map(|k| {
                        o.properties.get(&k).map(|v| {
                            let pair = vec![Value::String(Arc::from(k.as_str())), v.clone()];
                            Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                        })
                    })
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    // getOwnPropertyNames — like keys but includes non-enumerable
    // (our model doesn't track enumerability; alias to keys for MVP)
    vm.register_host_fn("wasm:js-object", "getOwnPropertyNames",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let keys: Vec<Value> = own_keys(&o).into_iter()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    // getOwnPropertySymbols — we don't distinguish symbol vs string keys yet
    vm.register_host_fn("wasm:js-object", "getOwnPropertySymbols",
        Box::new(|_ctx, _args| {
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-object", "length",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::I32(own_keys(&o).len() as i32);
            }
            Value::I32(0)
        }));
}

// ── Property descriptors ──────────────────────────────────────────────

fn register_descriptors(vm: &mut VM) {
    // defineProperty(obj, key, descriptor) -> obj
    // Descriptor is itself an object with {value, writable, enumerable,
    // configurable} or {get, set, enumerable, configurable} fields.
    // MVP: just extract `value` and do a plain set.
    vm.register_host_fn("wasm:js-object", "defineProperty",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let val = match args.get(2) {
                    Some(Value::Object(desc)) => {
                        let d = desc.lock().unwrap();
                        d.properties.get("value").cloned().unwrap_or(Value::Undefined)
                    }
                    _ => Value::Undefined,
                };
                {
                    let mut o = obj.lock().unwrap();
                    o.properties.insert(key, val);
                }
                return Value::Object(obj);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-object", "defineProperties",
        Box::new(|_ctx, args| {
            if let (Some(target), Some(Value::Object(descs))) = (obj_of(args, 0), args.get(1)) {
                let d = descs.lock().unwrap();
                let entries: Vec<(String, Value)> = d.properties.iter()
                    .filter(|(k, _)| !k.starts_with("__"))
                    .filter_map(|(k, v)| {
                        if let Value::Object(dv) = v {
                            let dlock = dv.lock().unwrap();
                            dlock.properties.get("value").cloned().map(|val| (k.clone(), val))
                        } else {
                            None
                        }
                    })
                    .collect();
                drop(d);
                let mut t = target.lock().unwrap();
                for (k, v) in entries {
                    t.properties.insert(k, v);
                }
                drop(t);
                return Value::Object(target);
            }
            Value::Null
        }));

    // getOwnPropertyDescriptor(obj, key) -> descriptor or undefined
    vm.register_host_fn("wasm:js-object", "getOwnPropertyDescriptor",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                if let Some(v) = o.properties.get(&key) {
                    let mut desc = Object::new();
                    desc.properties.insert("value".into(), v.clone());
                    desc.properties.insert("writable".into(), Value::I32(1));
                    desc.properties.insert("enumerable".into(), Value::I32(1));
                    desc.properties.insert("configurable".into(), Value::I32(1));
                    return Value::Object(Arc::new(Mutex::new(desc)));
                }
            }
            Value::Undefined
        }));

    // getOwnPropertyDescriptors(obj) -> { key: descriptor, ... }
    vm.register_host_fn("wasm:js-object", "getOwnPropertyDescriptors",
        Box::new(|_ctx, args| {
            let mut result = Object::new();
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                for (k, v) in &o.properties {
                    if k.starts_with("__") { continue; }
                    let mut desc = Object::new();
                    desc.properties.insert("value".into(), v.clone());
                    desc.properties.insert("writable".into(), Value::I32(1));
                    desc.properties.insert("enumerable".into(), Value::I32(1));
                    desc.properties.insert("configurable".into(), Value::I32(1));
                    result.properties.insert(k.clone(), Value::Object(Arc::new(Mutex::new(desc))));
                }
            }
            Value::Object(Arc::new(Mutex::new(result)))
        }));
}

// ── Prototype ─────────────────────────────────────────────────────────

fn register_prototype(vm: &mut VM) {
    vm.register_host_fn("wasm:js-object", "getPrototypeOf",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return o.properties.get(PROTO_KEY).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-object", "setPrototypeOf",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let proto = args.get(1).cloned().unwrap_or(Value::Null);
                let mut o = obj.lock().unwrap();
                o.properties.insert(PROTO_KEY.into(), proto);
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }));
}

// ── Locking (freeze / seal / preventExtensions) ───────────────────────

fn register_locking(vm: &mut VM) {
    vm.register_host_fn("wasm:js-object", "freeze",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut o = obj.lock().unwrap();
                o.properties.insert(FROZEN_MARK.into(), Value::I32(1));
                o.properties.insert(SEALED_MARK.into(), Value::I32(1));
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-object", "isFrozen",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::I32(if o.properties.get(FROZEN_MARK).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-object", "seal",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut o = obj.lock().unwrap();
                o.properties.insert(SEALED_MARK.into(), Value::I32(1));
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-object", "isSealed",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::I32(if o.properties.get(SEALED_MARK).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-object", "preventExtensions",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut o = obj.lock().unwrap();
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-object", "isExtensible",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::I32(match o.properties.get(EXTENSIBLE_MARK) {
                    Some(Value::I32(0)) => 0,
                    _ => 1,  // absence => extensible
                });
            }
            Value::I32(0)
        }));
}

// ── Comparison ────────────────────────────────────────────────────────

fn register_comparison(vm: &mut VM) {
    // Object.is(a, b) — SameValue: NaN === NaN, -0 distinct from +0
    vm.register_host_fn("wasm:js-object", "is",
        Box::new(|_ctx, args| {
            let a = args.first();
            let b = args.get(1);
            let same = match (a, b) {
                (Some(Value::F64(x)), Some(Value::F64(y))) => {
                    if x.is_nan() && y.is_nan() {
                        true
                    } else if *x == 0.0 && *y == 0.0 {
                        // Distinguish +0 vs -0 via sign bit
                        x.is_sign_positive() == y.is_sign_positive()
                    } else {
                        x == y
                    }
                }
                (Some(x), Some(y)) => x.eq(y),
                _ => false,
            };
            Value::I32(if same { 1 } else { 0 })
        }));
}

// ── Prototype methods (called via obj.foo()) ──────────────────────────

fn register_prototype_methods(vm: &mut VM) {
    vm.register_host_fn("wasm:js-object", "hasOwnProperty",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                return Value::I32(if o.properties.contains_key(&key) { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-object", "isPrototypeOf",
        Box::new(|_ctx, args| {
            // isPrototypeOf: is `self` in `other`'s prototype chain?
            if let (Some(self_obj), Some(other)) = (obj_of(args, 0), obj_of(args, 1)) {
                let mut current = other;
                loop {
                    let o = current.lock().unwrap();
                    let proto = o.properties.get(PROTO_KEY).cloned();
                    drop(o);
                    match proto {
                        Some(Value::Object(p)) => {
                            if Arc::ptr_eq(&p, &self_obj) {
                                return Value::I32(1);
                            }
                            if Arc::ptr_eq(&p, &current) {
                                return Value::I32(0);
                            }
                            current = p;
                        }
                        _ => return Value::I32(0),
                    }
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-object", "propertyIsEnumerable",
        Box::new(|_ctx, args| {
            // Our model: any own property is enumerable.
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                return Value::I32(if o.properties.contains_key(&key) && !key.starts_with("__") { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    // toString(): spec default is "[object Object]" for plain objects
    vm.register_host_fn("wasm:js-object", "toString",
        Box::new(|_ctx, args| {
            if is_object(args.first().unwrap_or(&Value::Null)) {
                return Value::String(Arc::from("[object Object]"));
            }
            Value::String(Arc::from(""))
        }));

    vm.register_host_fn("wasm:js-object", "toLocaleString",
        Box::new(|_ctx, args| {
            if is_object(args.first().unwrap_or(&Value::Null)) {
                return Value::String(Arc::from("[object Object]"));
            }
            Value::String(Arc::from(""))
        }));

    // valueOf: spec default returns the object itself
    vm.register_host_fn("wasm:js-object", "valueOf",
        Box::new(|_ctx, args| args.first().cloned().unwrap_or(Value::Null)));
}

// ── PHP extensions ────────────────────────────────────────────────────

fn register_php_extensions(vm: &mut VM) {
    // appendAutoKey(obj, value) -> i32 key
    // Implements PHP's `$a[] = x` — finds the next int key (max of
    // existing int keys + 1, or 0 if none) and sets it.
    vm.register_host_fn("wasm:js-object", "appendAutoKey",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let val = args.get(1).cloned().unwrap_or(Value::Null);
                let mut o = obj.lock().unwrap();
                // Get and increment the counter, initializing if absent.
                let next = match o.properties.get(NEXT_INT_KEY) {
                    Some(Value::I32(n)) => *n,
                    Some(Value::F64(n)) => *n as i32,
                    _ => {
                        // Compute from existing int-keyed entries
                        let mut max_k: i32 = -1;
                        for k in o.properties.keys() {
                            if let Ok(n) = k.parse::<i32>() {
                                if n > max_k { max_k = n; }
                            }
                        }
                        max_k + 1
                    }
                };
                o.properties.insert(next.to_string(), val);
                o.properties.insert(NEXT_INT_KEY.into(), Value::I32(next + 1));
                return Value::I32(next);
            }
            Value::I32(0)
        }));
}
