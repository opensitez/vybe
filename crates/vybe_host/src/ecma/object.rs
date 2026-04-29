//! # `ecma:object` host handlers
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
use vybe_bytecode::VM;

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
    vm.register_host_fn("ecma:object", "new",
        Box::new(|_ctx, _args| {
            Value::Object(Arc::new(Mutex::new(Object::new())))
        }));

    // create(proto) -> new obj with prototype link
    vm.register_host_fn("ecma:object", "create",
        Box::new(|_ctx, args| {
            // Object.create(proto) — ECMA-262 §20.1.2.2.
            //
            // True spec semantics ([[Get]] walking [[Prototype]]) need
            // the JS compiler to emit `ecma:object:get(obj, key)` for
            // property access — currently `obj.foo` lowers to STRUCT_GET
            // which does own-only lookup. Until that migration lands,
            // copy parent's enumerable own properties down so STRUCT_GET
            // finds inherited members; also stash the parent under
            // `__proto__` so reflective ops like `getPrototypeOf` work.
            // Internal `__`-prefixed metadata is skipped during copy.
            let mut obj = Object::new();
            match args.first() {
                Some(proto @ Value::Object(p)) => {
                    obj.properties.insert(PROTO_KEY.into(), proto.clone());
                    let parent = p.lock().unwrap();
                    for (k, v) in parent.properties.iter() {
                        if k.starts_with("__") { continue; }
                        obj.properties.insert(k.clone(), v.clone());
                    }
                }
                Some(Value::Null) => {
                    obj.properties.insert(PROTO_KEY.into(), Value::Null);
                }
                _ => {}
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // fromEntries(iterable) -> new obj
    vm.register_host_fn("ecma:object", "fromEntries",
        Box::new(|_ctx, args| {
            let mut obj = Object::new();
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                match s.kind {
                    ObjectKind::Array(ref pairs) => {
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
                    // Map iterates as `[key, value]` pairs (§24.1.3.5).
                    ObjectKind::Map(ref m) => {
                        for (k, v) in m.iter() {
                            obj.properties.insert(key_string(k), v.clone());
                        }
                    }
                    _ => {}
                }
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }));

    // `Object.assign(target, ...sources)` — ECMA-262 §20.1.2.1.
    // Variadic in the source positions; each source contributes its
    // own enumerable string-keyed properties onto target. Returns the
    // modified target. Internal `__`-prefixed properties are skipped
    // (they're our private metadata, not enumerable JS properties).
    vm.register_host_fn("ecma:object", "assign",
        Box::new(|_ctx, args| {
            let target = match args.first() {
                Some(t) => t.clone(),
                None => return Value::Null,
            };
            if let Value::Object(t) = &target {
                for source in args.iter().skip(1) {
                    if let Value::Object(s) = source {
                        let props: Vec<(String, Value)> = {
                            let src = s.lock().unwrap();
                            src.properties.iter()
                                .filter(|(k, _)| !k.starts_with("__"))
                                .map(|(k, v)| (k.clone(), v.clone()))
                                .collect()
                        };
                        let mut tgt = t.lock().unwrap();
                        for (k, v) in props {
                            tgt.properties.insert(k, v);
                        }
                    }
                }
            }
            target
        }));
}

// ── Property access ───────────────────────────────────────────────────

fn register_access(vm: &mut VM) {
    // get(obj, key) -> value (walks prototype chain)
    vm.register_host_fn("ecma:object", "get",
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
    vm.register_host_fn("ecma:object", "set",
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
    vm.register_host_fn("ecma:object", "has",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                return Value::Bool(proto_walk_get(&obj, &key).is_some());
            }
            Value::Bool(false)
        }));

    // hasOwn(obj, key) -> bool (own-only, no prototype walk). Polymorphic
    // over Array / Map / Ordinary. Backs JS `Object.hasOwn` + `in`
    // operator, PHP `array_key_exists`, Python `key in dict`, Ruby
    // `Hash#key?`. Returns Value::Bool so string coercion gives
    // "true"/"false" (ECMA-262 §23.1.2.3).
    vm.register_host_fn("ecma:object", "hasOwn",
        Box::new(|_ctx, args| {
            let key_raw = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                let found = match &o.kind {
                    ObjectKind::Array(v) => {
                        let i = key_raw.as_i32();
                        i >= 0 && (i as usize) < v.len()
                    }
                    ObjectKind::Map(m) => {
                        if m.contains_key(&key_raw) { true }
                        else if let Value::String(s) = &key_raw {
                            s.parse::<i32>().ok().map_or(false, |n| m.contains_key(&Value::I32(n)))
                        } else if let Value::I32(n) = &key_raw {
                            m.contains_key(&Value::String(Arc::from(n.to_string().as_str())))
                        } else { false }
                    }
                    _ => {
                        let key = args.get(1).map(key_string).unwrap_or_default();
                        o.properties.contains_key(&key)
                    }
                };
                return Value::Bool(found);
            }
            Value::Bool(false)
        }));

    // delete(obj, key) -> bool — ECMA-262 §13.5.1 (delete operator).
    vm.register_host_fn("ecma:object", "delete",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let mut o = obj.lock().unwrap();
                if o.properties.get(SEALED_MARK).is_some() {
                    return Value::Bool(false);
                }
                return Value::Bool(o.properties.remove(&key).is_some());
            }
            Value::Bool(false)
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

    // Polymorphic over Array / Map / Ordinary. Portable: scripts compiled
    // against `ecma:object.keys` run on any WASM engine (V8,
    // SpiderMonkey, wasmtime with the js-object polyfill). Every language
    // (PHP `array_keys`, Python `dict.keys`, Ruby `Hash#keys`, JS
    // `Object.keys`) binds to this SAME import.
    // Helper: return the ordered string keys of an Ordinary object.
    // Honors the `__keys` tracker (set by dict::emit_new + friends)
    // for JS-spec insertion-order semantics. Falls back to own_keys
    // order when no tracker is present (legacy / C# / VB class
    // instances that don't use the tracker).
    fn ordinary_ordered_keys(o: &Object) -> Vec<String> {
        if let Some(Value::Object(keys_arr)) = o.properties.get("__keys") {
            let ka = keys_arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ka.kind {
                return elems.iter()
                    .filter_map(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
                    .filter(|k| o.properties.contains_key(k))
                    .collect();
            }
        }
        own_keys(o)
    }

    vm.register_host_fn("ecma:object", "keys",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let keys: Vec<Value> = (0..v.len())
                            .map(|i| Value::String(Arc::from(i.to_string().as_str())))
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                    }
                    ObjectKind::Map(m) => {
                        let keys: Vec<Value> = m.keys().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                    }
                    _ => {}
                }
                let keys: Vec<Value> = ordinary_ordered_keys(&o).into_iter()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("ecma:object", "values",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(v.clone()))));
                    }
                    ObjectKind::Map(m) => {
                        let vals: Vec<Value> = m.values().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
                    }
                    _ => {}
                }
                let values: Vec<Value> = ordinary_ordered_keys(&o).into_iter()
                    .filter_map(|k| o.properties.get(&k).cloned())
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(values))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("ecma:object", "entries",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => {
                        let entries: Vec<Value> = v.iter().enumerate()
                            .map(|(i, val)| {
                                let pair = vec![Value::I32(i as i32), val.clone()];
                                Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                    }
                    ObjectKind::Map(m) => {
                        let entries: Vec<Value> = m.iter()
                            .map(|(k, v)| {
                                let pair = vec![k.clone(), v.clone()];
                                Value::Object(Arc::new(Mutex::new(Object::new_array(pair))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                    }
                    _ => {}
                }
                let entries: Vec<Value> = ordinary_ordered_keys(&o).into_iter()
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
    vm.register_host_fn("ecma:object", "getOwnPropertyNames",
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
    vm.register_host_fn("ecma:object", "getOwnPropertySymbols",
        Box::new(|_ctx, _args| {
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("ecma:object", "length",
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
    vm.register_host_fn("ecma:object", "defineProperty",
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

    vm.register_host_fn("ecma:object", "defineProperties",
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
    vm.register_host_fn("ecma:object", "getOwnPropertyDescriptor",
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
    vm.register_host_fn("ecma:object", "getOwnPropertyDescriptors",
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
    vm.register_host_fn("ecma:object", "getPrototypeOf",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return o.properties.get(PROTO_KEY).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:object", "setPrototypeOf",
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
    vm.register_host_fn("ecma:object", "freeze",
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

    vm.register_host_fn("ecma:object", "isFrozen",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(o.properties.get(FROZEN_MARK).is_some());
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:object", "seal",
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

    vm.register_host_fn("ecma:object", "isSealed",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(o.properties.get(SEALED_MARK).is_some());
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:object", "preventExtensions",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let mut o = obj.lock().unwrap();
                o.properties.insert(EXTENSIBLE_MARK.into(), Value::I32(0));
                drop(o);
                return Value::Object(obj);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:object", "isExtensible",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let o = obj.lock().unwrap();
                return Value::Bool(!matches!(o.properties.get(EXTENSIBLE_MARK), Some(Value::I32(0))));
            }
            Value::Bool(false)
        }));
}

// ── Comparison ────────────────────────────────────────────────────────

fn register_comparison(vm: &mut VM) {
    // Object.is(a, b) — SameValue: NaN === NaN, -0 distinct from +0
    vm.register_host_fn("ecma:object", "is",
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
            // ECMA-262 §20.1.2.13: returns a Boolean.
            Value::Bool(same)
        }));
}

// ── Prototype methods (called via obj.foo()) ──────────────────────────

fn register_prototype_methods(vm: &mut VM) {
    vm.register_host_fn("ecma:object", "hasOwnProperty",
        Box::new(|_ctx, args| {
            if let Some(obj) = obj_of(args, 0) {
                let key = args.get(1).map(key_string).unwrap_or_default();
                let o = obj.lock().unwrap();
                return Value::I32(if o.properties.contains_key(&key) { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("ecma:object", "isPrototypeOf",
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

    vm.register_host_fn("ecma:object", "propertyIsEnumerable",
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
    vm.register_host_fn("ecma:object", "toString",
        Box::new(|_ctx, args| {
            if is_object(args.first().unwrap_or(&Value::Null)) {
                return Value::String(Arc::from("[object Object]"));
            }
            Value::String(Arc::from(""))
        }));

    vm.register_host_fn("ecma:object", "toLocaleString",
        Box::new(|_ctx, args| {
            if is_object(args.first().unwrap_or(&Value::Null)) {
                return Value::String(Arc::from("[object Object]"));
            }
            Value::String(Arc::from(""))
        }));

    // valueOf: spec default returns the object itself
    vm.register_host_fn("ecma:object", "valueOf",
        Box::new(|_ctx, args| args.first().cloned().unwrap_or(Value::Null)));
}

// ── PHP extensions ────────────────────────────────────────────────────

fn register_php_extensions(vm: &mut VM) {
    // appendAutoKey(obj, value) -> i32 key
    // Implements PHP's `$a[] = x` — finds the next int key (max of
    // existing int keys + 1, or 0 if none) and sets it.
    vm.register_host_fn("ecma:object", "appendAutoKey",
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
