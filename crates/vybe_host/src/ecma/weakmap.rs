//! # `ecma:weakmap` and `ecma:weakset` host handlers
//!
//! Native Rust impls of `WeakMap.*` / `WeakSet.*` per ECMA-262 §24.3 /
//! §24.4.
//!
//! ## Weak reference caveat
//!
//! WASM GC MVP doesn't yet have weak references (it's Post-MVP). On
//! Vybe VM, for true JS-compat semantics we'd use `weak_table`'s
//! `WeakKeyHashMap` over `Arc<Object>`. For the MVP handler we use
//! strong references — functionally correct (get/set/has/delete all
//! behave per spec) but entries live as long as the WeakMap does,
//! rather than being collected when the key's external references
//! are gone.
//!
//! Phase B4 upgrade will swap in real weak references. Until then:
//! code that depends on strong-cleanup semantics (garbage-collected
//! caches, etc.) will retain more memory than on v8. Functional
//! correctness is unaffected.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

const WEAKMAP_TAG: &str = "__vybe_js_weakmap";
const WEAKSET_TAG: &str = "__vybe_js_weakset";
const WM_KEYS_PROP: &str = "__vybe_wm_keys";
// Values live in the backing Array (ObjectKind::Array); keys live in a
// parallel Array in the properties bag.

fn new_weakmap() -> Value {
    let mut obj = Object::new_array(Vec::new());
    obj.properties.insert(WEAKMAP_TAG.into(), Value::I32(1));
    obj.properties.insert(WM_KEYS_PROP.into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new())))));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn new_weakset() -> Value {
    let mut obj = Object::new_array(Vec::new());
    obj.properties.insert(WEAKSET_TAG.into(), Value::I32(1));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_weakmap(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(WEAKMAP_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn is_weakset(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(WEAKSET_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

/// WeakMap/WeakSet keys must be objects (spec trap on primitive keys).
/// We match keys by Arc pointer identity, not value equality.
fn key_ptr_find(keys: &[Value], key: &Value) -> Option<usize> {
    if let Value::Object(key_arc) = key {
        for (i, k) in keys.iter().enumerate() {
            if let Value::Object(existing) = k {
                if Arc::ptr_eq(existing, key_arc) {
                    return Some(i);
                }
            }
        }
    }
    None
}

pub fn register(vm: &mut VM) {
    register_weakmap(vm);
    register_weakset(vm);
}

// ── WeakMap ───────────────────────────────────────────────────────────

fn register_weakmap(vm: &mut VM) {
    vm.register_host_fn("ecma:weakmap", "new",
        Box::new(|_ctx, _args| new_weakmap()));

    vm.register_host_fn("ecma:weakmap", "fromIterable",
        Box::new(|_ctx, args| {
            let m = new_weakmap();
            if let Value::Object(mapobj) = &m {
                if let Some(Value::Object(src)) = args.first() {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref pairs) = s.kind {
                        let pairs = pairs.clone();
                        drop(s);
                        for pair in pairs {
                            if let Value::Object(p) = pair {
                                let pl = p.lock().unwrap();
                                if let ObjectKind::Array(ref kv) = pl.kind {
                                    if kv.len() >= 2 {
                                        weakmap_set(mapobj, kv[0].clone(), kv[1].clone());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            m
        }));

    vm.register_host_fn("ecma:weakmap", "get",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_weakmap(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                if !matches!(key, Value::Object(_)) {
                    return Value::Undefined;
                }
                let m = mapobj.lock().unwrap();
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref keys) = ko.kind {
                        if let Some(pos) = key_ptr_find(keys, &key) {
                            drop(ko);
                            if let ObjectKind::Array(ref values) = m.kind {
                                return values.get(pos).cloned().unwrap_or(Value::Undefined);
                            }
                        }
                    }
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("ecma:weakmap", "set",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_weakmap(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                // Per spec: non-object keys throw TypeError. MVP returns the map without inserting.
                if !matches!(key, Value::Object(_)) {
                    return Value::Object(mapobj);
                }
                weakmap_set(&mapobj, key, val);
                return Value::Object(mapobj);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:weakmap", "has",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_weakmap(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                if !matches!(key, Value::Object(_)) { return Value::I32(0); }
                let m = mapobj.lock().unwrap();
                if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref keys) = ko.kind {
                        return Value::I32(if key_ptr_find(keys, &key).is_some() { 1 } else { 0 });
                    }
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("ecma:weakmap", "delete",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_weakmap(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                if !matches!(key, Value::Object(_)) { return Value::I32(0); }
                let mut m = mapobj.lock().unwrap();
                let pos = if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref keys) = ko.kind {
                        key_ptr_find(keys, &key)
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
                    return Value::I32(1);
                }
            }
            Value::I32(0)
        }));
}

fn weakmap_set(mapobj: &Arc<Mutex<Object>>, key: Value, val: Value) {
    let mut m = mapobj.lock().unwrap();
    let existing = if let Some(Value::Object(keys_obj)) = m.properties.get(WM_KEYS_PROP) {
        let ko = keys_obj.lock().unwrap();
        if let ObjectKind::Array(ref keys) = ko.kind {
            key_ptr_find(keys, &key)
        } else {
            None
        }
    } else {
        None
    };
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
}

// ── WeakSet ───────────────────────────────────────────────────────────

fn register_weakset(vm: &mut VM) {
    vm.register_host_fn("ecma:weakset", "new",
        Box::new(|_ctx, _args| new_weakset()));

    vm.register_host_fn("ecma:weakset", "fromIterable",
        Box::new(|_ctx, args| {
            let s = new_weakset();
            if let Value::Object(setobj) = &s {
                if let Some(Value::Object(src)) = args.first() {
                    let srclock = src.lock().unwrap();
                    if let ObjectKind::Array(ref items) = srclock.kind {
                        let items = items.clone();
                        drop(srclock);
                        let mut so = setobj.lock().unwrap();
                        for item in items {
                            if !matches!(item, Value::Object(_)) { continue; }
                            if let ObjectKind::Array(ref vs) = so.kind {
                                if key_ptr_find(vs, &item).is_some() { continue; }
                            }
                            if let ObjectKind::Array(ref mut vs) = so.kind {
                                vs.push(item);
                            }
                        }
                    }
                }
            }
            s
        }));

    vm.register_host_fn("ecma:weakset", "add",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_weakset(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                if !matches!(v, Value::Object(_)) { return Value::Object(setobj); }
                let mut so = setobj.lock().unwrap();
                if let ObjectKind::Array(ref vs) = so.kind {
                    if key_ptr_find(vs, &v).is_some() {
                        drop(so);
                        return Value::Object(setobj);
                    }
                }
                if let ObjectKind::Array(ref mut vs) = so.kind {
                    vs.push(v);
                }
                drop(so);
                return Value::Object(setobj);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:weakset", "has",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_weakset(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                if !matches!(v, Value::Object(_)) { return Value::I32(0); }
                let so = setobj.lock().unwrap();
                if let ObjectKind::Array(ref vs) = so.kind {
                    return Value::I32(if key_ptr_find(vs, &v).is_some() { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("ecma:weakset", "delete",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_weakset(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                if !matches!(v, Value::Object(_)) { return Value::I32(0); }
                let mut so = setobj.lock().unwrap();
                let pos = if let ObjectKind::Array(ref vs) = so.kind {
                    key_ptr_find(vs, &v)
                } else {
                    None
                };
                if let Some(pos) = pos {
                    if let ObjectKind::Array(ref mut vs) = so.kind {
                        vs.remove(pos);
                    }
                    return Value::I32(1);
                }
            }
            Value::I32(0)
        }));
}
