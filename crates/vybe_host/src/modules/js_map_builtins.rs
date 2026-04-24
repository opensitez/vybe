//! # `vybe:js-map` and `vybe:js-set` host handlers
//!
//! Native Rust impls of `Map.prototype.*` / `Set.prototype.*` per
//! ECMA-262 §24.1 / §24.2, satisfying the imports declared in
//! `crates/vybe_bytecode/src/wasm/js_map_builtins.rs` and
//! `js_set_builtins.rs`.
//!
//! ## Storage (Phase B4)
//!
//! Map — `ObjectKind::Map(IndexMap<Value, Value>)`.
//! O(1) average-case get/set/has/delete while preserving JS-spec
//! insertion order for iteration. Keys use `SameValueZero` semantics
//! via `Value`'s `Hash + Eq` impls (NaN === NaN, -0 === +0,
//! integer-equal numerics collapse to the same key regardless of
//! `I32` / `I64` / `F64` source type).
//!
//! Set — still using the tagged-property-bag backing for MVP; a
//! dedicated `ObjectKind::Set(IndexSet<Value>)` lands in a follow-up
//! B4 pass alongside the ArrayBuffer / DataView / TypedArray variants.
//!
//! ## Behavioral contract
//!
//! Phase B6 behavioral tests in
//! `crates/vybe_host/tests/js_builtins_behavior_test.rs` lock down:
//!   - Map set/get roundtrip with string keys
//!   - Map has/delete correctness
//!   - Map size tracks insertions
//!   - Identity preservation of externref values through set + get
//! All tests continue to pass across this rewrite.
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use indexmap::IndexMap;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

// ── Map: variant-backed O(1) impl ─────────────────────────────────────

fn new_map_value() -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(IndexMap::new());
    obj.properties.insert("size".into(), Value::I32(0));
    Value::Object(Arc::new(Mutex::new(obj)))
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

// ── Set: variant-backed O(1) impl ─────────────────────────────────────

fn new_set() -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    obj.properties.insert("size".into(), Value::I32(0));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn is_set(args: &[Value], idx: usize) -> Option<Arc<Mutex<Object>>> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if matches!(o.kind, ObjectKind::Set(_)) {
            drop(o);
            return Some(obj.clone());
        }
    }
    None
}

fn sync_set_size(obj: &mut Object) {
    if let ObjectKind::Set(ref s) = obj.kind {
        let n = s.len() as i32;
        obj.properties.insert("size".into(), Value::I32(n));
    }
}

pub fn register(vm: &mut VM) {
    register_map(vm);
    register_set(vm);
}

// ── vybe:js-map ────────────────────────────────────────────────────────

fn register_map(vm: &mut VM) {
    vm.register_host_fn("vybe:js-map", "new",
        Box::new(|_ctx, _args| new_map_value()));

    // fromEntries(iterable) — iterable is an Array of [k, v] pairs.
    vm.register_host_fn("vybe:js-map", "fromEntries",
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

    vm.register_host_fn("vybe:js-map", "get",
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

    vm.register_host_fn("vybe:js-map", "set",
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

    vm.register_host_fn("vybe:js-map", "has",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::I32(if im.contains_key(&key) { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-map", "delete",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let key = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut m = mapobj.lock().unwrap();
                let removed = if let ObjectKind::Map(ref mut im) = m.kind {
                    // `shift_remove` preserves insertion order of the
                    // remaining entries (matches ECMA-262 §24.1.3.3
                    // "removes the element with key P and returns
                    // true if the element was present").
                    im.shift_remove(&key).is_some()
                } else {
                    false
                };
                sync_map_size(&mut m);
                return Value::I32(if removed { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-map", "clear",
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

    vm.register_host_fn("vybe:js-map", "size",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    return Value::I32(im.len() as i32);
                }
            }
            Value::I32(0)
        }));

    // keys / values / entries — Array snapshots in insertion order
    vm.register_host_fn("vybe:js-map", "keys",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let keys: Vec<Value> = im.keys().cloned().collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("vybe:js-map", "values",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let vals: Vec<Value> = im.values().cloned().collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    vm.register_host_fn("vybe:js-map", "entries",
        Box::new(|_ctx, args| {
            if let Some(mapobj) = is_map(args, 0) {
                let m = mapobj.lock().unwrap();
                if let ObjectKind::Map(ref im) = m.kind {
                    let pairs: Vec<Value> = im.iter()
                        .map(|(k, v)| Value::Object(Arc::new(Mutex::new(
                            Object::new_array(vec![k.clone(), v.clone()])
                        ))))
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    // forEach(map, callback) — invokes callback(value, key, map) per
    // entry in insertion order. ECMA-262 §24.1.3.5.
    vm.register_host_fn("vybe:js-map", "forEach",
        Box::new(|ctx, args| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
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
                    ctx.invoke(&callback, &invoke_args);
                }
            }
            Value::Undefined
        }));

    // Map.groupBy(iterable, fn) → Map — groups iterable entries by the
    // value fn returns for each. ES2025.
    vm.register_host_fn("vybe:js-map", "groupBy",
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
                    let invoke_args = vec![item.clone(), Value::I32(i as i32)];
                    let key = ctx.invoke(&callback, &invoke_args);
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
}

// ── vybe:js-set ────────────────────────────────────────────────────────

fn register_set(vm: &mut VM) {
    vm.register_host_fn("vybe:js-set", "new",
        Box::new(|_ctx, _args| new_set()));

    vm.register_host_fn("vybe:js-set", "fromIterable",
        Box::new(|_ctx, args| {
            let s = new_set();
            if let Value::Object(setobj) = &s {
                if let Some(Value::Object(src)) = args.first() {
                    let srclock = src.lock().unwrap();
                    if let ObjectKind::Array(ref items) = srclock.kind {
                        let items = items.clone();
                        drop(srclock);
                        let mut so = setobj.lock().unwrap();
                        if let ObjectKind::Set(ref mut s) = so.kind {
                            for item in items { s.insert(item); }
                        }
                        sync_set_size(&mut so);
                    }
                }
            }
            s
        }));

    vm.register_host_fn("vybe:js-set", "add",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                {
                    let mut so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref mut s) = so.kind {
                        s.insert(v);
                    }
                    sync_set_size(&mut so);
                }
                return Value::Object(setobj);
            }
            Value::Null
        }));

    vm.register_host_fn("vybe:js-set", "has",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    return Value::I32(if s.contains(&v) { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-set", "delete",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                let mut so = setobj.lock().unwrap();
                let removed = if let ObjectKind::Set(ref mut s) = so.kind {
                    // `shift_remove` preserves insertion order of the
                    // remaining members per ECMA-262 §24.2.3.4.
                    s.shift_remove(&v)
                } else {
                    false
                };
                sync_set_size(&mut so);
                return Value::I32(if removed { 1 } else { 0 });
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-set", "clear",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let mut so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref mut s) = so.kind { s.clear(); }
                sync_set_size(&mut so);
            }
            Value::Null
        }));

    vm.register_host_fn("vybe:js-set", "size",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    return Value::I32(s.len() as i32);
                }
            }
            Value::I32(0)
        }));

    for name in &["values", "keys"] {
        vm.register_host_fn("vybe:js-set", name,
            Box::new(|_ctx, args| {
                if let Some(setobj) = is_set(args, 0) {
                    let so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref s) = so.kind {
                        let snapshot: Vec<Value> = s.iter().cloned().collect();
                        return Value::Object(Arc::new(Mutex::new(Object::new_array(snapshot))));
                    }
                }
                Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
            }));
    }

    vm.register_host_fn("vybe:js-set", "entries",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    let pairs: Vec<Value> = s.iter()
                        .map(|v| Value::Object(Arc::new(Mutex::new(
                            Object::new_array(vec![v.clone(), v.clone()])
                        ))))
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))))
        }));

    // Set.prototype.forEach(callback) — callback receives (value,
    // value, set) — the key mirrors the value per §24.2.3.6.
    vm.register_host_fn("vybe:js-set", "forEach",
        Box::new(|ctx, args| {
            let callback = args.get(1).cloned().unwrap_or(Value::Null);
            if let Some(setobj) = is_set(args, 0) {
                let snapshot: Vec<Value> = {
                    let so = setobj.lock().unwrap();
                    if let ObjectKind::Set(ref s) = so.kind {
                        s.iter().cloned().collect()
                    } else {
                        Vec::new()
                    }
                };
                for v in snapshot {
                    let invoke_args = vec![
                        v.clone(), v, Value::Object(setobj.clone()),
                    ];
                    ctx.invoke(&callback, &invoke_args);
                }
            }
            Value::Undefined
        }));

    // ── Set algebra (ES2025) ────────────────────────────────────────
    //
    // IndexSet gives us native `.union` / `.intersection` methods, but
    // we hand-roll here to preserve ECMA-262's insertion-order
    // semantics for the result: "iterate a first, then take b's
    // members that aren't in a" for `union`; "iterate a, keep those
    // also in b" for `intersection`; etc.

    vm.register_host_fn("vybe:js-set", "union",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let Value::Object(outobj) = &out {
                let mut o = outobj.lock().unwrap();
                if let ObjectKind::Set(ref mut os) = o.kind {
                    for arg_idx in 0..2 {
                        if let Some(setobj) = is_set(args, arg_idx) {
                            let so = setobj.lock().unwrap();
                            if let ObjectKind::Set(ref s) = so.kind {
                                for v in s.iter() { os.insert(v.clone()); }
                            }
                        }
                    }
                }
                sync_set_size(&mut o);
            }
            out
        }));

    vm.register_host_fn("vybe:js-set", "intersection",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                if let Value::Object(outobj) = &out {
                    let alock = a.lock().unwrap();
                    let block = b.lock().unwrap();
                    let mut o = outobj.lock().unwrap();
                    if let (ObjectKind::Set(avs), ObjectKind::Set(bvs),
                            ObjectKind::Set(out_s))
                        = (&alock.kind, &block.kind, &mut o.kind)
                    {
                        for v in avs.iter() {
                            if bvs.contains(v) {
                                out_s.insert(v.clone());
                            }
                        }
                    }
                    sync_set_size(&mut o);
                }
            }
            out
        }));

    vm.register_host_fn("vybe:js-set", "difference",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                if let Value::Object(outobj) = &out {
                    let alock = a.lock().unwrap();
                    let block = b.lock().unwrap();
                    let mut o = outobj.lock().unwrap();
                    if let (ObjectKind::Set(avs), ObjectKind::Set(bvs),
                            ObjectKind::Set(out_s))
                        = (&alock.kind, &block.kind, &mut o.kind)
                    {
                        for v in avs.iter() {
                            if !bvs.contains(v) {
                                out_s.insert(v.clone());
                            }
                        }
                    }
                    sync_set_size(&mut o);
                }
            }
            out
        }));

    vm.register_host_fn("vybe:js-set", "symmetricDifference",
        Box::new(|_ctx, args| {
            let out = new_set();
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                if let Value::Object(outobj) = &out {
                    let alock = a.lock().unwrap();
                    let block = b.lock().unwrap();
                    let mut o = outobj.lock().unwrap();
                    if let (ObjectKind::Set(avs), ObjectKind::Set(bvs),
                            ObjectKind::Set(out_s))
                        = (&alock.kind, &block.kind, &mut o.kind)
                    {
                        for v in avs.iter() {
                            if !bvs.contains(v) {
                                out_s.insert(v.clone());
                            }
                        }
                        for v in bvs.iter() {
                            if !avs.contains(v) {
                                out_s.insert(v.clone());
                            }
                        }
                    }
                    sync_set_size(&mut o);
                }
            }
            out
        }));

    vm.register_host_fn("vybe:js-set", "isSubsetOf",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Set(avs), ObjectKind::Set(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let is_sub = avs.iter().all(|v| bvs.contains(v));
                    return Value::I32(if is_sub { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-set", "isSupersetOf",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Set(avs), ObjectKind::Set(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let is_super = bvs.iter().all(|v| avs.contains(v));
                    return Value::I32(if is_super { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));

    vm.register_host_fn("vybe:js-set", "isDisjointFrom",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Set(avs), ObjectKind::Set(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let disjoint = !avs.iter().any(|v| bvs.contains(v));
                    return Value::I32(if disjoint { 1 } else { 0 });
                }
            }
            Value::I32(0)
        }));
}
