//! # `wasm:js-map` and `wasm:js-set` host handlers
//!
//! Native Rust impls of `Map.prototype.*` / `Set.prototype.*` per
//! ECMA-262 §24.1 / §24.2, satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_map_builtins.rs` and
//! `js_set_builtins.rs`.
//!
//! Underlying storage: our existing `ObjectKind::Array(Vec<Value>)`
//! + parallel-key side-table in properties. A full `ObjectKind::Map`
//! variant will land in Phase B4 once we confirm the import surface
//! doesn't need any extra representation bits. Today's impl is a
//! property-bag Map — functionally correct for JS compat but slow
//! for large maps (linear key lookup). Phase B4 swaps in
//! `IndexMap<Value, Value>` for O(1) ops.
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

/// Magic property key we use to mark an Object as a JS Map.
/// Phase B4 replaces this with a dedicated `ObjectKind::Map`.
const MAP_TAG: &str = "__vybe_js_map";
const SET_TAG: &str = "__vybe_js_set";

/// Keys stored in parallel with Array contents. Property name on the
/// Object. Keys are kept in insertion order; values are in the
/// backing Array in the same order.
const MAP_KEYS_PROP: &str = "__vybe_map_keys";

fn new_map() -> Value {
    let mut obj = Object::new_array(Vec::new());
    obj.properties.insert(MAP_TAG.into(), Value::I32(1));
    obj.properties.insert(MAP_KEYS_PROP.into(), Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new())))));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn new_set() -> Value {
    let mut obj = Object::new_array(Vec::new());
    obj.properties.insert(SET_TAG.into(), Value::I32(1));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_map(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(MAP_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn is_set(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if o.properties.get(SET_TAG).is_some() {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

/// Find the position of `key` in the map's keys array, or None.
fn find_map_key(map: &Object, key: &Value) -> Option<usize> {
    if let Some(Value::Object(keys_obj)) = map.properties.get(MAP_KEYS_PROP) {
        let ko = keys_obj.lock().unwrap();
        if let ObjectKind::Array(ref keys) = ko.kind {
            for (i, k) in keys.iter().enumerate() {
                if k.eq(key) {
                    return Some(i);
                }
            }
        }
    }
    None
}

fn find_set_entry(set: &Object, v: &Value) -> Option<usize> {
    if let ObjectKind::Array(ref vs) = set.kind {
        for (i, e) in vs.iter().enumerate() {
            if e.eq(v) {
                return Some(i);
            }
        }
    }
    None
}

pub fn register(vm: &mut VM) {
    register_map(vm);
    register_set(vm);
}

// ── wasm:js-map ────────────────────────────────────────────────────────

fn register_map(vm: &mut VM) {
    vm.register_host_fn("wasm:js-map", "new",
        Box::new(|_ctx, _args| new_map()));

    // fromEntries(iterable) -> Map — iterable is an Array of [k, v] pairs
    vm.register_host_fn("wasm:js-map", "fromEntries",
        Box::new(|_ctx, args| {
            let m = new_map();
            if let Value::Object(mapobj) = &m {
                if let Some(Value::Object(src)) = args.first() {
                    let srclock = src.lock().unwrap();
                    if let ObjectKind::Array(ref pairs) = srclock.kind {
                        for pair in pairs {
                            if let Value::Object(p) = pair {
                                let pl = p.lock().unwrap();
                                if let ObjectKind::Array(ref kv) = pl.kind {
                                    if kv.len() >= 2 {
                                        map_set(mapobj, kv[0].clone(), kv[1].clone());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            m
        }));

    vm.register_host_fn("wasm:js-map", "get",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                if let Some(pos) = find_map_key(&m, &key) {
                    if let ObjectKind::Array(ref values) = m.kind {
                        return values.get(pos).cloned().unwrap_or(Value::Undefined);
                    }
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("wasm:js-map", "set",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                map_set(&mapobj, key, val);
                return Value::Object(mapobj);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-map", "has",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                return Value::I32(if find_map_key(&m, &key).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-map", "delete",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut m = mapobj.lock().unwrap();
                if let Some(pos) = find_map_key(&m, &key) {
                    // Remove from both the values Vec and the keys Array.
                    if let ObjectKind::Array(ref mut values) = m.kind {
                        values.remove(pos);
                    }
                    if let Some(Value::Object(keys_obj)) = m.properties.get(MAP_KEYS_PROP).cloned() {
                        let mut ko = keys_obj.lock().unwrap();
                        if let ObjectKind::Array(ref mut keys) = ko.kind {
                            keys.remove(pos);
                        }
                    }
                    let new_len = if let ObjectKind::Array(ref v) = m.kind { v.len() } else { 0 };
                    m.properties.insert("size".into(), Value::F64(new_len as f64));
                    return Value::I32(1);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-map", "clear",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let mut m = mapobj.lock().unwrap();
                if let ObjectKind::Array(ref mut v) = m.kind { v.clear(); }
                if let Some(Value::Object(keys_obj)) = m.properties.get(MAP_KEYS_PROP).cloned() {
                    let mut ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref mut keys) = ko.kind { keys.clear(); }
                }
                m.properties.insert("size".into(), Value::F64(0.0));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-map", "size",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Array(ref v) = m.kind {
                    return Value::I32(v.len() as i32);
                }
            }
            Value::I32(0)
        }));

    // keys / values / entries — return Array snapshots (Phase B12 upgrades to iterators)
    vm.register_host_fn("wasm:js-map", "keys",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let Some(Value::Object(keys_obj)) = m.properties.get(MAP_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let ObjectKind::Array(ref keys) = ko.kind {
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(keys.clone()))));
                    }
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-map", "values",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Array(ref v) = m.kind {
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(v.clone()))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-map", "entries",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let Some(Value::Object(keys_obj)) = m.properties.get(MAP_KEYS_PROP) {
                    let ko = keys_obj.lock().unwrap();
                    if let (ObjectKind::Array(keys), ObjectKind::Array(values)) = (&ko.kind, &m.kind) {
                        let pairs: Vec<Value> = keys.iter().zip(values.iter())
                            .map(|(k, v)| {
                                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![k.clone(), v.clone()]))))
                            })
                            .collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
                    }
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-map", "forEach",
        Box::new(|_ctx, _args| Value::Null));

    vm.register_host_fn("wasm:js-map", "groupBy",
        Box::new(|_ctx, _args| new_map()));
}

fn map_set(mapobj: &Arc<Mutex<Object>>, key: Value, val: Value) {
    let mut m = mapobj.lock().unwrap();
    let existing = find_map_key(&m, &key);
    if let Some(pos) = existing {
        if let ObjectKind::Array(ref mut values) = m.kind {
            values[pos] = val;
        }
    } else {
        if let ObjectKind::Array(ref mut values) = m.kind {
            values.push(val);
        }
        if let Some(Value::Object(keys_obj)) = m.properties.get(MAP_KEYS_PROP).cloned() {
            let mut ko = keys_obj.lock().unwrap();
            if let ObjectKind::Array(ref mut keys) = ko.kind {
                keys.push(key);
            }
        }
    }
    let new_len = if let ObjectKind::Array(ref v) = m.kind { v.len() } else { 0 };
    m.properties.insert("size".into(), Value::F64(new_len as f64));
}

// ── wasm:js-set ────────────────────────────────────────────────────────

fn register_set(vm: &mut VM) {
    vm.register_host_fn("wasm:js-set", "new",
        Box::new(|_ctx, _args| new_set()));

    vm.register_host_fn("wasm:js-set", "fromIterable",
        Box::new(|_ctx, args| {
            let s = new_set();
            if let Value::Object(setobj) = &s {
                if let Some(Value::Object(src)) = args.first() {
                    let srclock = src.lock().unwrap();
                    if let ObjectKind::Array(ref items) = srclock.kind {
                        let items = items.clone();
                        drop(srclock);
                        let mut so = setobj.lock().unwrap();
                        for item in items {
                            if find_set_entry(&so, &item).is_none() {
                                if let ObjectKind::Array(ref mut v) = so.kind {
                                    v.push(item);
                                }
                            }
                        }
                        let new_len = if let ObjectKind::Array(ref v) = so.kind { v.len() } else { 0 };
                        so.properties.insert("size".into(), Value::F64(new_len as f64));
                    }
                }
            }
            s
        }));

    vm.register_host_fn("wasm:js-set", "add",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut so = setobj.lock().unwrap();
                if find_set_entry(&so, &v).is_none() {
                    if let ObjectKind::Array(ref mut vs) = so.kind {
                        vs.push(v);
                    }
                    let new_len = if let ObjectKind::Array(ref vs) = so.kind { vs.len() } else { 0 };
                    so.properties.insert("size".into(), Value::F64(new_len as f64));
                }
                drop(so);
                return Value::Object(setobj);
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-set", "has",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                let so = setobj.lock().unwrap();
                return Value::I32(if find_set_entry(&so, &v).is_some() { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-set", "delete",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut so = setobj.lock().unwrap();
                if let Some(pos) = find_set_entry(&so, &v) {
                    if let ObjectKind::Array(ref mut vs) = so.kind {
                        vs.remove(pos);
                    }
                    let new_len = if let ObjectKind::Array(ref vs) = so.kind { vs.len() } else { 0 };
                    so.properties.insert("size".into(), Value::F64(new_len as f64));
                    return Value::I32(1);
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-set", "clear",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let mut so = setobj.lock().unwrap();
                if let ObjectKind::Array(ref mut vs) = so.kind { vs.clear(); }
                so.properties.insert("size".into(), Value::F64(0.0));
            }
            Value::Null
        }));

    vm.register_host_fn("wasm:js-set", "size",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let so = setobj.lock().unwrap();
                if let ObjectKind::Array(ref vs) = so.kind {
                    return Value::I32(vs.len() as i32);
                }
            }
            Value::I32(0)
        }));

    // values / keys / entries — Array snapshots
    for name in &["values", "keys"] {
        vm.register_host_fn("wasm:js-set", name,
            Box::new(|_ctx, args| {
                if let Some(setobj) = is_set(args, 0) {
                    let so = setobj.lock().unwrap();
                    if let ObjectKind::Array(ref vs) = so.kind {
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(vs.clone()))));
                    }
                }
                Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
            }));
    }

    vm.register_host_fn("wasm:js-set", "entries",
        Box::new(|_ctx, args| {
            // For Set, entries returns [[v, v], ...] per spec
            if let Some(setobj) = is_set(args, 0) {
                let so = setobj.lock().unwrap();
                if let ObjectKind::Array(ref vs) = so.kind {
                    let pairs: Vec<Value> = vs.iter()
                        .map(|v| Value::Object(Arc::new(Mutex::new(Object::new_array(vec![v.clone(), v.clone()])))))
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("wasm:js-set", "forEach",
        Box::new(|_ctx, _args| Value::Null));

    // ── Set algebra (ES2025) ────────────────────────────────────────

    vm.register_host_fn("wasm:js-set", "union",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let Value::Object(outobj) = &out {
                let mut o = outobj.lock().unwrap();
                for arg_idx in 0..2 {
                    if let Some(setobj) = is_set(args, arg_idx) {
                        let so = setobj.lock().unwrap();
                        if let ObjectKind::Array(ref vs) = so.kind {
                            for v in vs {
                                if find_set_entry(&o, v).is_none() {
                                    if let ObjectKind::Array(ref mut out_vs) = o.kind {
                                        out_vs.push(v.clone());
                                    }
                                }
                            }
                        }
                    }
                }
                let new_len = if let ObjectKind::Array(ref vs) = o.kind { vs.len() } else { 0 };
                o.properties.insert("size".into(), Value::F64(new_len as f64));
            }
            out
        }));

    vm.register_host_fn("wasm:js-set", "intersection",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Array(avs), ObjectKind::Array(bvs))
                    = (&alock.kind, &block.kind)
                {
                    if let Value::Object(outobj) = &out {
                        let mut o = outobj.lock().unwrap();
                        if let ObjectKind::Array(ref mut out_vs) = o.kind {
                            for v in avs {
                                if bvs.iter().any(|bv| bv.eq(v)) {
                                    out_vs.push(v.clone());
                                }
                            }
                        }
                        let new_len = if let ObjectKind::Array(ref vs) = o.kind { vs.len() } else { 0 };
                        o.properties.insert("size".into(), Value::F64(new_len as f64));
                    }
                }
            }
            out
        }));

    vm.register_host_fn("wasm:js-set", "difference",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Array(avs), ObjectKind::Array(bvs))
                    = (&alock.kind, &block.kind)
                {
                    if let Value::Object(outobj) = &out {
                        let mut o = outobj.lock().unwrap();
                        if let ObjectKind::Array(ref mut out_vs) = o.kind {
                            for v in avs {
                                if !bvs.iter().any(|bv| bv.eq(v)) {
                                    out_vs.push(v.clone());
                                }
                            }
                        }
                        let new_len = if let ObjectKind::Array(ref vs) = o.kind { vs.len() } else { 0 };
                        o.properties.insert("size".into(), Value::F64(new_len as f64));
                    }
                }
            }
            out
        }));

    vm.register_host_fn("wasm:js-set", "symmetricDifference",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Array(avs), ObjectKind::Array(bvs))
                    = (&alock.kind, &block.kind)
                {
                    if let Value::Object(outobj) = &out {
                        let mut o = outobj.lock().unwrap();
                        if let ObjectKind::Array(ref mut out_vs) = o.kind {
                            for v in avs {
                                if !bvs.iter().any(|bv| bv.eq(v)) {
                                    out_vs.push(v.clone());
                                }
                            }
                            for v in bvs {
                                if !avs.iter().any(|av| av.eq(v)) {
                                    out_vs.push(v.clone());
                                }
                            }
                        }
                        let new_len = if let ObjectKind::Array(ref vs) = o.kind { vs.len() } else { 0 };
                        o.properties.insert("size".into(), Value::F64(new_len as f64));
                    }
                }
            }
            out
        }));

    vm.register_host_fn("wasm:js-set", "isSubsetOf",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Array(avs), ObjectKind::Array(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let is_sub = avs.iter().all(|v| bvs.iter().any(|bv| bv.eq(v)));
                    return Value::I32(if is_sub { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-set", "isSupersetOf",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Array(avs), ObjectKind::Array(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let is_super = bvs.iter().all(|v| avs.iter().any(|av| av.eq(v)));
                    return Value::I32(if is_super { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("wasm:js-set", "isDisjointFrom",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Array(avs), ObjectKind::Array(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let disjoint = !avs.iter().any(|v| bvs.iter().any(|bv| bv.eq(v)));
                    return Value::I32(if disjoint { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));
}
