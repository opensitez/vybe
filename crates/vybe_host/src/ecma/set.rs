//! # `ecma:set` — ECMA-262 §24.2 Set
//!
//! Native Rust impls of `Set.prototype.*` + ES2025 set-algebra methods
//! (`union`, `intersection`, `difference`, `symmetricDifference`,
//! `isSubsetOf`, `isSupersetOf`, `isDisjointFrom`).
//!
//! Backing storage is `ObjectKind::Set(IndexSet<Value>)` — O(1) avg
//! add/has/delete while preserving insertion order for iteration.
//! Membership uses `SameValueZero` via `Value`'s `Hash + Eq`.
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

fn new_set() -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    obj.properties.insert("size".into(), Value::I32(0));
    // __type stamp: see comment on `ecma:map.new`. Without it the
    // TypeRegistry-driven `STRUCT_GET s "add"` lookup misses.
    obj.properties.insert("__type".into(), Value::String(Arc::from("Set")));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn new_set_from_iterable(args: &[Value]) -> Value {
    let s = new_set();
    if let (Value::Object(setobj), Some(Value::Object(src))) = (&s, args.first()) {
        let srclock = src.lock().unwrap();
        if let ObjectKind::Array(ref items) = srclock.kind {
            let items = items.clone();
            drop(srclock);
            let mut so = setobj.lock().unwrap();
            if let ObjectKind::Set(ref mut iset) = so.kind {
                for item in items {
                    iset.insert(item);
                }
            }
            sync_set_size(&mut so);
        }
    }
    s
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
    // `new Set(iterable?)` — per ECMA-262 §24.2.1.1 the constructor optionally
    // takes an iterable whose elements become Set members.
    vm.register_host_fn("ecma:set", "new",
        Box::new(|_ctx, args| new_set_from_iterable(args)));

    vm.register_host_fn("ecma:set", "fromIterable",
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

    vm.register_host_fn("ecma:set", "add",
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

    vm.register_host_fn("ecma:set", "has",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let v = args.get(1).cloned().unwrap_or(Value::Undefined);
                let so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref s) = so.kind {
                    return Value::Bool(s.contains(&v));
                }
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:set", "delete",
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
                return Value::Bool(removed);
            }
            Value::Bool(false)
        }));

    vm.register_host_fn("ecma:set", "clear",
        Box::new(|_ctx, args| {
            if let Some(setobj) = is_set(args, 0) {
                let mut so = setobj.lock().unwrap();
                if let ObjectKind::Set(ref mut s) = so.kind { s.clear(); }
                sync_set_size(&mut so);
            }
            Value::Null
        }));

    vm.register_host_fn("ecma:set", "size",
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
        vm.register_host_fn("ecma:set", name,
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

    vm.register_host_fn("ecma:set", "entries",
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
    vm.register_host_fn("ecma:set", "forEach",
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

    vm.register_host_fn("ecma:set", "union",
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

    vm.register_host_fn("ecma:set", "intersection",
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

    vm.register_host_fn("ecma:set", "difference",
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

    vm.register_host_fn("ecma:set", "symmetricDifference",
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

    vm.register_host_fn("ecma:set", "isSubsetOf",
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

    vm.register_host_fn("ecma:set", "isSupersetOf",
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

    vm.register_host_fn("ecma:set", "isDisjointFrom",
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

    // .NET HashSet mutating set algebra — `UnionWith` / `IntersectWith` /
    // `ExceptWith` / `SymmetricExceptWith` modify the receiver in place.
    // Distinct from the immutable ES2025 `union` / `intersection` / etc.
    // which return a fresh Set. The ES variants are still registered above;
    // these mutate variants are the .NET-shape entry points.
    vm.register_host_fn("ecma:set", "unionWith",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let to_add: Vec<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.iter().cloned().collect()
                    } else {
                        Vec::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    for v in to_add { avs.insert(v); }
                }
                sync_set_size(&mut alock);
            }
            Value::Undefined
        }));

    vm.register_host_fn("ecma:set", "intersectWith",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let b_snapshot: Vec<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.iter().cloned().collect()
                    } else {
                        Vec::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    avs.retain(|v| b_snapshot.contains(v));
                }
                sync_set_size(&mut alock);
            }
            Value::Undefined
        }));

    vm.register_host_fn("ecma:set", "exceptWith",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let b_snapshot: Vec<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.iter().cloned().collect()
                    } else {
                        Vec::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    avs.retain(|v| !b_snapshot.contains(v));
                }
                sync_set_size(&mut alock);
            }
            Value::Undefined
        }));

    vm.register_host_fn("ecma:set", "symmetricExceptWith",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let b_snapshot: Vec<Value> = {
                    let block = b.lock().unwrap();
                    if let ObjectKind::Set(ref bvs) = block.kind {
                        bvs.iter().cloned().collect()
                    } else {
                        Vec::new()
                    }
                };
                let mut alock = a.lock().unwrap();
                if let ObjectKind::Set(ref mut avs) = alock.kind {
                    let mut to_remove = Vec::new();
                    let mut to_add = Vec::new();
                    for v in &b_snapshot {
                        if avs.contains(v) {
                            to_remove.push(v.clone());
                        } else {
                            to_add.push(v.clone());
                        }
                    }
                    avs.retain(|v| !to_remove.contains(v));
                    for v in to_add { avs.insert(v); }
                }
                sync_set_size(&mut alock);
            }
            Value::Undefined
        }));

    vm.register_host_fn("ecma:set", "overlaps",
        Box::new(|_ctx, args| {
            if let (Some(a), Some(b)) = (is_set(args, 0), is_set(args, 1)) {
                let alock = a.lock().unwrap();
                let block = b.lock().unwrap();
                if let (ObjectKind::Set(avs), ObjectKind::Set(bvs))
                    = (&alock.kind, &block.kind)
                {
                    let overlap = avs.iter().any(|v| bvs.contains(v));
                    return Value::Bool(overlap);
                }
            }
            Value::Bool(false)
        }));
}
