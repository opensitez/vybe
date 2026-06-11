//! Built-in .NET types: DateTime, StringBuilder, List, Dictionary.
//! Each constructor creates an object with methods as HostFunctions.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // `register_datetime` retired — DateTime constructor + statics
    // lower through `emitter::dotnet::core::datetime_adapter` to
    // `ecma:date.{new, now, parse}` (which read
    // `wasi:clocks/wall-clock.now`).
    // `register_timespan` retired — TimeSpan statics lower through
    // `emitter::dotnet::core::timespan_adapter` (pure inline
    // bytecode, no host fns).
    register_list(vm);
    register_dictionary(vm);
    register_queue_stack(vm);
}

// ============================================================
// DateTime
// ============================================================

// `register_datetime` retired — DateTime constructor + statics
// (`Now` / `UtcNow` / `Today` / `Parse`) lower at compile time
// through `emitter::dotnet::core::datetime_adapter` to
// `ecma:date.{new, now, parse}` host fns (which read through
// `wasi:clocks/wall-clock.now`). The `make_datetime_*` helpers
// are gone too — adapter logic is bytecode, not host-Rust.
fn register_list(vm: &mut VM) {
    // .NET `List<T>` runtime dispatch lowers entirely through the
    // dotnet wrapper Component Model adapter (see
    // `emitter/dotnet/core/component_classes.rs::List`) — methods
    // route to `collections.*` (Op::ARRAY_NEW + ecma:array.*). Only
    // the few range-shaped operations that don't have a 1:1 ECMA
    // equivalent stay here as host-fn primitives.
    //
    // Live primitives:
    //   listNew              — constructor (stamps `__type=List`)
    //   listInsertRange     — `List.InsertRange(index, collection)`
    //   listRemoveRange     — `List.RemoveRange(index, count)`
    //   listGetRange        — `List.GetRange(index, count)` → new List
    //   listSetRange        — `List.SetRange(index, collection)`
    //   listBinarySearch    — `List.BinarySearch(value)` → index | -1
    //
    // Everything else (`listAdd` / `listRemove` / `listSort` / …) was
    // retired when the dotnet wrapper switched to `collections.*`.

    vm.register_host_fn("vybe:types", "listNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:types", "listInsertRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(dst)), Some(Value::Object(src))) = (args.first(), args.get(2)) {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let s = src.lock().unwrap();
            if let ObjectKind::Array(ref src_elems) = s.kind {
                let items: Vec<Value> = src_elems.clone();
                drop(s);
                let mut d = dst.lock().unwrap();
                if let ObjectKind::Array(ref mut dst_elems) = d.kind {
                    let pos = idx.min(dst_elems.len());
                    for (i, item) in items.into_iter().enumerate() {
                        dst_elems.insert(pos + i, item);
                    }
                    let len = dst_elems.len() as f64;
                    d.properties.insert("length".into(), Value::F64(len));
                    d.properties.insert("count".into(), Value::F64(len));
                }
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:types", "listRemoveRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let count = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &mut o.kind {
                let start = idx.min(elems.len());
                let end = (start + count).min(elems.len());
                elems.drain(start..end);
                let len = elems.len() as f64;
                o.properties.insert("length".into(), Value::F64(len));
                o.properties.insert("count".into(), Value::F64(len));
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:types", "listGetRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let count = args.get(2).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let start = idx.min(elems.len());
                let end = (start + count).min(elems.len());
                let sub: Vec<Value> = elems[start..end].to_vec();
                let mut result = Object::new_array(sub);
                result.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(result)));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    vm.register_host_fn("vybe:types", "listSetRange", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(dst)), Some(Value::Object(src))) = (args.first(), args.get(2)) {
            let idx = args.get(1).map(|v| v.as_i32().max(0) as usize).unwrap_or(0);
            let s = src.lock().unwrap();
            if let ObjectKind::Array(ref src_elems) = s.kind {
                let items: Vec<Value> = src_elems.clone();
                drop(s);
                let mut d = dst.lock().unwrap();
                if let ObjectKind::Array(ref mut dst_elems) = d.kind {
                    for (i, item) in items.into_iter().enumerate() {
                        let pos = idx + i;
                        if pos < dst_elems.len() {
                            dst_elems[pos] = item;
                        }
                    }
                }
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:types", "listBinarySearch", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                for (i, e) in elems.iter().enumerate() {
                    if format!("{}", e) == search { return Value::F64(i as f64); }
                }
            }
        }
        Value::F64(-1.0)
    }));
}

// ============================================================
// Dictionary(Of K, V)
// ============================================================

fn register_dictionary(vm: &mut VM) {
    // .NET `Dictionary<K,V>` runtime dispatch is done at compile time
    // by the dotnet wrapper Component Model adapter (`Dictionary` →
    // `ecma:map.*`). Only the constructor + a couple of legacy
    // emitter-fallback helpers stay here.
    //
    // Live primitives:
    //   dictNew    — bare constructor (stamps `__type=Dictionary`)
    //   dictAdd    — used by `compiler_common::dict::emit_set` as a
    //                polymorphic stack-based set (handles both `__data`
    //                property-bag dicts and direct property writes).
    //   dictItem   — same shape as `dictAdd`, polymorphic get.
    //   dictValues — polymorphic values collection.
    //
    // Everything else (`dictRemove`, `dictKeys`, `dictClear`,
    // `dictContainsKey`, `dictTryGetValue`) was retired — those calls
    // now hit `ecma:map.*` directly through the dotnet adapter.

    vm.register_host_fn("vybe:types", "dictNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Dictionary")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:types", "dictAdd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(2).cloned().unwrap_or(Value::Null);
            let data = {
                let o = obj.lock().unwrap();
                match o.properties.get("__data") {
                    Some(Value::Object(data)) => Some(data.clone()),
                    _ => None,
                }
            };
            if let Some(data) = data {
                let count = {
                    let mut data_obj = data.lock().unwrap();
                    data_obj.properties.insert(key.clone(), value.clone());
                    data_obj.properties.len() as f64
                };
                let mut outer = obj.lock().unwrap();
                if !key.starts_with("__") {
                    outer.properties.insert(key, value);
                }
                outer.properties.insert("count".into(), Value::F64(count));
                outer.properties.insert("length".into(), Value::F64(count));
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:types", "dictItem", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            if let Some(val) = o.properties.get(&key) {
                return val.clone();
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:types", "dictValues", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let vals: Vec<Value> = data.lock().unwrap().properties.values().cloned().collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
            let vals: Vec<Value> = o.properties.iter()
                .filter(|(k, _)| !k.starts_with("__"))
                .map(|(_, v)| v.clone())
                .collect();
            if !vals.is_empty() {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));
}
// ============================================================
// Queue / Stack / HashSet
// ============================================================

fn register_queue_stack(vm: &mut VM) {
    // .NET `Queue<T>` / `Stack<T>` / `HashSet<T>` runtime dispatch is
    // done at compile time by the dotnet wrapper Component Model
    // adapter (`Queue` / `Stack` route to `collections.*`, `HashSet`
    // routes to `ecma:set.*`). Only the constructors stay here as
    // host primitives to stamp `__type`.

    vm.register_host_fn("vybe:types", "queueNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("Queue")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:types", "stackNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new_array(vec![]);
        obj.properties.insert("__type".into(), Value::String(Arc::from("Stack")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:types", "hashSetNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object {
            properties: std::collections::HashMap::new(),
            kind: ObjectKind::Set(indexmap::IndexSet::new()),
            type_id: 0,
            fields: Vec::new(),
        };
        obj.properties.insert("__type".into(), Value::String(Arc::from("HashSet")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));
}
// `register_timespan` body deleted — see `pub fn register` for the
// retirement note. All TimeSpan factory statics lower through the
// emitter::dotnet::core::timespan_adapter (pure inline bytecode).

// ============================================================
// Guid
// ============================================================

// `register_guid` retired — `.NET Guid` is a UUID v4 string per
// RFC 4122. `Guid.NewGuid()` namespace binding now points at
// `wasi:random/random.uuid` (WASI 0.2.8 spec primitive); `Guid.Parse`
// is a string passthrough via `ecma:string.String`. No `vybe:types`
// involvement.

// ============================================================
// Primitive type statics (Double, Single, Boolean, Decimal)
// ============================================================

// `register_primitives` retired entirely — `Double` / `Boolean` /
// `Array` static method primitives all migrated:
//   - bool/double parsers were dead code (no callers).
//   - `Array.Clear/Copy/Resize/Sort` lower at compile time through
//     `emitter::dotnet::core::array_adapter` to bundled stdlib
//     bytecode chunks (`__vybe_sort_in_place`, `__vybe_array_copy`,
//     `__vybe_redim`) composing `ecma:array.*` primitives.
