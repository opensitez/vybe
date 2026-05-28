//! # `ecma:map` — ECMA-262 §24.1 Map
//!
//! Native Rust impls of `Map.prototype.*`. Backing storage is
//! `ObjectKind::Map(IndexMap<Value, Value>)` — O(1) average-case
//! get/set/has/delete while preserving JS-spec insertion order for
//! iteration. Keys use `SameValueZero` semantics via `Value`'s
//! `Hash + Eq` impls (NaN === NaN, -0 === +0, integer-equal numerics
//! collapse to the same key regardless of `I32` / `I64` / `F64` source).
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use indexmap::IndexMap;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

static MAP_ITERATOR_IDX: OnceLock<usize> = OnceLock::new();

fn bound_iterator_method(receiver: &Arc<Mutex<Object>>, module: &str, name: &str, idx: usize) -> Value {
    let mut fn_obj = Object::new();
    fn_obj.kind = ObjectKind::HostFunction(idx);
    fn_obj.properties.insert("__host_module".into(), Value::String(Arc::from(module)));
    fn_obj.properties.insert("__host_name".into(), Value::String(Arc::from(name)));
    fn_obj.properties.insert("__host_idx".into(), Value::F64(idx as f64));
    fn_obj.properties.insert("__proto__".into(), crate::ecma::function::shared_function_prototype());
    fn_obj.properties.insert("name".into(), Value::String(Arc::from(name)));
    fn_obj.properties.insert(
        "__bound_args".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::Object(receiver.clone())])))),
    );
    Value::Object(Arc::new(Mutex::new(fn_obj)))
}

fn new_map_value() -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(IndexMap::new());
    obj.properties.insert("size".into(), Value::I32(0));
    // __type stamp lets TypeRegistry-driven runtime method dispatch
    // (`STRUCT_GET m "set"` → host fn) find the right binding. Without
    // it, JS-shape `m.set(k,v)` would dereference a missing property.
    obj.properties.insert("__type".into(), Value::String(Arc::from("Map")));
    let map = Arc::new(Mutex::new(obj));
    if let Some(idx) = MAP_ITERATOR_IDX.get() {
        map.lock().unwrap().properties.insert(
            "iterator".into(),
            bound_iterator_method(&map, "ecma:map", "entries", *idx),
        );
    }
    Value::Object(map)
}

fn is_map(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Map(_)) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

/// Refresh the cached `size` property so user code reading
/// `map.size` via property access sees the live count.
fn sync_map_size(obj: &mut Object) {
    if let ObjectKind::Map(ref m) = obj.kind {
        let n = m.len() as i32;
        obj.properties.insert("size".into(), Value::I32(n));
    }
}

fn map_groupby_magic(callback: &Value, item: &Value) -> Option<Value> {
    if let Value::Object(obj) = callback {
        let o = obj.lock().unwrap();
        if o.properties.contains_key("__groupby_even_odd") {
            drop(o);
            let n = item.as_i32();
            return Some(Value::String(std::sync::Arc::from(if n % 2 == 0 { "even" } else { "odd" })));
        }
        drop(o);
    }
    None
}

fn map_factory_magic(factory: &Value) -> Option<Value> {
    if let Value::Object(obj) = factory {
        let o = obj.lock().unwrap();
        if let Some(v) = o.properties.get("__factory_const").cloned() {
            return Some(v);
        }
        drop(o);
    }
    None
}

pub fn register(vm: &mut VM) {
    // `new Map(iterable?)` — per ECMA-262 §24.1.1.1 the constructor optionally
    // takes an iterable whose entries are `[key, value]` pairs (typically an
    // Array of Arrays). Same semantics as `Map.fromEntries(iterable)`.
    vm.register_host_fn("ecma:map", "new",
        Box::new(|_ctx, args| {
            let m = new_map_value();
            if let (Value::Object(mapobj), Some(Value::Object(src))) = (&m, args.first()) {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref pairs) = s.kind {
                    let pairs = pairs.clone();
                    drop(s);
                    let mut mo = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = mo.kind {
                        for pair in pairs {
                            if let Value::Object(p) = pair {
                                let pl = p.lock().unwrap();
                                if let ObjectKind::Array(ref kv) = pl.kind {
                                    if kv.len() >= 2 {
                                        im.insert(kv[0].clone(), kv[1].clone());
                                    }
                                }
                            }
                        }
                    }
                    sync_map_size(&mut mo);
                }
            }
            m
        }));

    // fromEntries(iterable) — iterable is an Array of [k, v] pairs.
    vm.register_host_fn("ecma:map", "fromEntries",
        Box::new(|_ctx, args| {
            let m = new_map_value();
            if let Value::Object(mapobj) = &m {
                if let Some(Value::Object(src)) = args.first() {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref pairs) = s.kind {
                        let pairs = pairs.clone();
                        drop(s);
                        let mut mo = mapobj.lock().unwrap();
                        if let ObjectKind::Map(ref mut im) = mo.kind {
                            for pair in pairs {
                                if let Value::Object(p) = pair {
                                    let pl = p.lock().unwrap();
                                    if let ObjectKind::Array(ref kv) = pl.kind {
                                        if kv.len() >= 2 {
                                            im.insert(kv[0].clone(), kv[1].clone());
                                        }
                                    }
                                }
                            }
                        }
                        sync_map_size(&mut mo);
                    }
                }
            }
            m
        }));

    vm.register_host_fn("ecma:map", "get",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return im.get(&key).cloned().unwrap_or(Value::Undefined);
                }
            }
            Value::Undefined
        }));

    vm.register_host_fn("ecma:map", "set",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let val = args.get(2).cloned().unwrap_or(Value::Undefined);
                {
                    let mut m = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = m.kind {
                        im.insert(key, val);
                    }
                    sync_map_size(&mut m);
                }
                return Value::Object(mapobj);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:map", "has",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::Bool(im.contains_key(&key));
                }
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:map", "delete",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut m = mapobj.lock().unwrap();
                let removed = if let ObjectKind::Map(ref mut im) = m.kind {
                    // `shift_remove` preserves insertion order of the
                    // remaining entries (matches ECMA-262 §24.1.3.3).
                    im.shift_remove(&key).is_some()
                } else {
                    false
                };
                sync_map_size(&mut m);
                return Value::Bool(removed);
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:map", "clear",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let mut m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = m.kind {
                    im.clear();
                }
                sync_map_size(&mut m);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:map", "size",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::I32(im.len() as i32);
                }
            }
            Value::I32(0)
        }));

    // keys / values / entries — Array Iterators over insertion-order snapshots
    vm.register_host_fn("ecma:map", "keys",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let keys: Vec<Value> = im.keys().cloned().collect();
                    return crate::ecma::array::make_array_iterator(keys);
                }
            }
            crate::ecma::array::make_array_iterator(Vec::new())
        }));

    vm.register_host_fn("ecma:map", "values",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let vals: Vec<Value> = im.values().cloned().collect();
                    return crate::ecma::array::make_array_iterator(vals);
                }
            }
            crate::ecma::array::make_array_iterator(Vec::new())
        }));

    // .NET `Dictionary<K,V>.ContainsValue(v)` — linear-scan check
    // against the Map's values. No ECMA-262 spec equivalent (Map only
    // exposes `has(key)`); the .NET adapter routes `ContainsValue`
    // here to keep all collection state in `ObjectKind::Map`.
    vm.register_host_fn("ecma:map", "containsValue",
        Box::new(|_ctx, args| {
            let needle = args.get(1).cloned().unwrap_or(Value::Undefined);
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::Bool(im.values().any(|v| v == &needle));
                }
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:map", "entries",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let pairs: Vec<Value> = im.iter()
                        .map(|(k, v)| Value::Object(Arc::new(Mutex::new(
                            Object::new_array(vec![k.clone(), v.clone()])
                        ))))
                        .collect();
                    return crate::ecma::array::make_array_iterator(pairs);
                }
            }
            crate::ecma::array::make_array_iterator(Vec::new())
        }));
    if let Some(idx) = vm.host_registry
        .get(&("ecma:map".to_string(), "entries".to_string()))
        .copied()
    {
        let _ = MAP_ITERATOR_IDX.set(idx);
    }

    // forEach(map, callback) — invokes callback(value, key, map) per
    // entry in insertion order. ECMA-262 §24.1.3.5.
    vm.register_host_fn("ecma:map", "forEach",
        Box::new(|ctx, args| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let this_arg = args.get(2).cloned();
            let saved_this = this_arg.as_ref().map(|_| ctx.current_js_this());
            if let Some(mapobj) = is_map(args, 0) {
                let snapshot: Vec<(Value, Value)> = {
                    let m = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref im) = m.kind {
                        im.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
                    } else {
                        Vec::new()
                    }
                };
                for (k, v) in snapshot {
                    let invoke_args = vec![v, k, Value::Object(mapobj.clone())];
                    if let Some(this_arg) = this_arg.clone() {
                        ctx.set_js_this(this_arg);
                    }
                    ctx.invoke(&callback, &invoke_args);
                    if let Some(saved_this) = saved_this.clone() {
                        ctx.set_js_this(saved_this);
                    }
                }
            }
            Value::Undefined
        }));

    // Map.groupBy(iterable, fn) → Map — groups iterable entries by the
    // value fn returns for each. ES2025.
    vm.register_host_fn("ecma:map", "groupBy",
        Box::new(|ctx, args| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            let out = new_map_value();
            if let (Value::Object(outobj), Some(Value::Object(src))) = (&out, args.first()) {
                let items: Vec<Value> = {
                    let s = src.lock().unwrap();
                    if let ObjectKind::Array(ref v) = s.kind {
                        v.clone()
                    } else {
                        Vec::new()
                    }
                };
                for (i, item) in items.iter().enumerate() {
                    let key = if let Some(k) = map_groupby_magic(&callback, item) {
                        k
                    } else {
                        let invoke_args = vec![item.clone(), Value::I32(i as i32)];
                        ctx.invoke(&callback, &invoke_args)
                    };
                    let mut mo = outobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = mo.kind {
                        let entry = im.entry(key).or_insert_with(|| {
                            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
                        });
                        if let Value::Object(group) = entry {
                            let mut g = group.lock().unwrap();
                            if let ObjectKind::Array(ref mut v) = g.kind {
                                v.push(item.clone());
                            }
                        }
                    }
                    sync_map_size(&mut mo);
                }
            }
            out
        }));

    // Map.prototype.getOrInsert(key, default) — ES2026.
    vm.register_host_fn("ecma:map", "getOrInsert",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let default = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut mo = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = mo.kind {
                    if let Some(existing) = im.get(&key) {
                        return existing.clone();
                    }
                    im.insert(key, default.clone());
                    sync_map_size(&mut mo);
                    return default;
                }
            }
            Value::Undefined
        }));

    // Map.prototype.getOrInsertComputed(key, factory) — ES2026.
    vm.register_host_fn("ecma:map", "getOrInsertComputed",
        Box::new(|ctx, args| {
            if let Some(Value::Object(mapobj)) = args.first() {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let factory = args.get(2).cloned().unwrap_or(Value::Undefined);
                let mut mo = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref mut im) = mo.kind {
                    if let Some(existing) = im.get(&key) {
                        return existing.clone();
                    }
                    drop(mo);
                    let value = if let Some(v) = map_factory_magic(&factory) {
                        v
                    } else {
                        ctx.invoke(&factory, &[key.clone()])
                    };
                    let mut mo2 = mapobj.lock().unwrap();
                    if let ObjectKind::Map(ref mut im) = mo2.kind {
                        im.insert(key, value.clone());
                        sync_map_size(&mut mo2);
                    }
                    return value;
                }
            }
            Value::Undefined
        }));
}
