use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // -- Map constructor: new Map() or new Map([[k,v], ...]) --
    // Always create a fresh object; if first arg is an array of [k,v] pairs, populate.
    vm.register_host_fn("vybe:collections", "Map", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let this = Value::Object(Arc::new(Mutex::new(Object::new())));
        if let Value::Object(obj) = &this {
            let mut o = obj.lock().unwrap();
            o.properties.insert("__type".into(), Value::String(Arc::from("Map")));
            let data = Value::Object(Arc::new(Mutex::new(Object::new())));
            o.properties.insert("__data".into(), data.clone());
            let mut count: f64 = 0.0;
            // Accept iterable of [key, value] pairs
            if let Some(Value::Object(entries)) = args.first() {
                let e = entries.lock().unwrap();
                if let ObjectKind::Array(ref items) = e.kind {
                    if let Value::Object(data_obj) = &data {
                        let mut d = data_obj.lock().unwrap();
                        for item in items {
                            if let Value::Object(pair) = item {
                                let p = pair.lock().unwrap();
                                if let ObjectKind::Array(ref kv) = p.kind {
                                    if kv.len() >= 2 {
                                        let k = format!("{}", kv[0]);
                                        d.properties.insert(k, kv[1].clone());
                                        count += 1.0;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            o.properties.insert("size".into(), Value::F64(count));
        }
        this
    }));

    // -- WeakMap constructor: new WeakMap() -- (alias for Map, no GC semantics)
    vm.register_host_fn("vybe:collections", "WeakMap", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let obj = Object::new();
        let this = Value::Object(Arc::new(Mutex::new(obj)));
        if let Value::Object(o_arc) = &this {
            let mut o = o_arc.lock().unwrap();
            o.properties.insert("__type".into(), Value::String(Arc::from("WeakMap")));
            o.properties.insert("__data".into(), Value::Object(Arc::new(Mutex::new(Object::new()))));
        }
        this
    }));

    // -- WeakSet constructor: new WeakSet() -- (alias for Set)
    vm.register_host_fn("vybe:collections", "WeakSet", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let obj = Object::new();
        let this = Value::Object(Arc::new(Mutex::new(obj)));
        if let Value::Object(o_arc) = &this {
            let mut o = o_arc.lock().unwrap();
            o.properties.insert("__type".into(), Value::String(Arc::from("WeakSet")));
            let mut items = Object::new();
            items.kind = ObjectKind::Array(Vec::new());
            o.properties.insert("__items".into(), Value::Object(Arc::new(Mutex::new(items))));
        }
        this
    }));

    // Map.prototype.set(key, value)
    vm.register_host_fn("vybe:collections", "mapSet", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(2).cloned().unwrap_or(Value::Null);
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                data.lock().unwrap().properties.insert(key, value);
                let size = data.lock().unwrap().properties.len() as f64;
                drop(o);
                obj.lock().unwrap().properties.insert("size".into(), Value::F64(size));
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Map.prototype.get(key)
    vm.register_host_fn("vybe:collections", "mapGet", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
        }
        Value::Null
    }));

    // Map.prototype.has(key)
    vm.register_host_fn("vybe:collections", "mapHas", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                return Value::Bool(data.lock().unwrap().properties.contains_key(&key));
            }
        }
        Value::Bool(false)
    }));

    // Map.prototype.delete(key)
    vm.register_host_fn("vybe:collections", "mapDelete", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let existed = data.lock().unwrap().properties.remove(&key).is_some();
                let size = data.lock().unwrap().properties.len() as f64;
                drop(o);
                obj.lock().unwrap().properties.insert("size".into(), Value::F64(size));
                return Value::Bool(existed);
            }
        }
        Value::Bool(false)
    }));

    // Map.prototype.keys()
    vm.register_host_fn("vybe:collections", "mapKeys", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let keys: Vec<Value> = data.lock().unwrap().properties.keys()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // Map.prototype.values()
    vm.register_host_fn("vybe:collections", "mapValues", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                let vals: Vec<Value> = data.lock().unwrap().properties.values().cloned().collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // -- Set constructor: new Set() --
    vm.register_host_fn("vybe:collections", "Set", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // Always create a fresh object; if first arg is an array, init with deduped items.
        let this = Value::Object(Arc::new(Mutex::new(Object::new())));
        if let Value::Object(obj) = &this {
            let mut o = obj.lock().unwrap();
            o.properties.insert("__type".into(), Value::String(Arc::from("Set")));
            let mut items: Vec<Value> = Vec::new();
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
            if let Some(Value::Object(src)) = args.first() {
                let s = src.lock().unwrap();
                if let ObjectKind::Array(ref arr) = s.kind {
                    for v in arr {
                        let key = format!("{}", v);
                        if seen.insert(key) {
                            items.push(v.clone());
                        }
                    }
                }
            }
            let len = items.len() as f64;
            o.properties.insert("__items".into(), Value::Object(Arc::new(Mutex::new(Object::new_array(items)))));
            o.properties.insert("size".into(), Value::F64(len));
        }
        this
    }));

    // Set.prototype.add(value)
    vm.register_host_fn("vybe:collections", "setAdd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let value = args.get(1).cloned().unwrap_or(Value::Null);
            let value_str = format!("{}", value);
            let o = obj.lock().unwrap();
            if let Some(Value::Object(items)) = o.properties.get("__items") {
                let mut items_ref = items.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = items_ref.kind {
                    let exists = elems.iter().any(|e| format!("{}", e) == value_str);
                    if !exists {
                        elems.push(value);
                        let len = elems.len() as f64;
                        drop(items_ref);
                        drop(o);
                        obj.lock().unwrap().properties.insert("size".into(), Value::F64(len));
                    }
                }
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Set.prototype.has(value)
    vm.register_host_fn("vybe:collections", "setHas", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let value_str = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(items)) = o.properties.get("__items") {
                let items_ref = items.lock().unwrap();
                if let ObjectKind::Array(ref elems) = items_ref.kind {
                    return Value::Bool(elems.iter().any(|e| format!("{}", e) == value_str));
                }
            }
        }
        Value::Bool(false)
    }));

    // Set.prototype.delete(value)
    vm.register_host_fn("vybe:collections", "setDelete", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let value_str = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let o = obj.lock().unwrap();
            if let Some(Value::Object(items)) = o.properties.get("__items") {
                let mut items_ref = items.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = items_ref.kind {
                    let before = elems.len();
                    elems.retain(|e| format!("{}", e) != value_str);
                    let removed = elems.len() < before;
                    let len = elems.len() as f64;
                    drop(items_ref);
                    drop(o);
                    obj.lock().unwrap().properties.insert("size".into(), Value::F64(len));
                    return Value::Bool(removed);
                }
            }
        }
        Value::Bool(false)
    }));

    // Map.prototype.clear() — drop all entries from __data and reset size.
    // The type registry in builtin_types.rs already advertises `mapClear`
    // as the host fn for `Map.clear`, so wiring it here completes the path.
    vm.register_host_fn("vybe:collections", "mapClear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(data)) = o.properties.get("__data") {
                data.lock().unwrap().properties.clear();
            }
            drop(o);
            obj.lock().unwrap().properties.insert("size".into(), Value::F64(0.0));
        }
        Value::Null
    }));

    // Set.prototype.clear() — drop all elements from __items and reset size.
    vm.register_host_fn("vybe:collections", "setClear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(items)) = o.properties.get("__items") {
                let mut items_ref = items.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = items_ref.kind {
                    elems.clear();
                }
            }
            drop(o);
            obj.lock().unwrap().properties.insert("size".into(), Value::F64(0.0));
        }
        Value::Null
    }));

    // Set.prototype.values()
    vm.register_host_fn("vybe:collections", "setValues", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(Value::Object(items)) = o.properties.get("__items") {
                let items_ref = items.lock().unwrap();
                if let ObjectKind::Array(ref elems) = items_ref.kind {
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
                }
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // -- Generic has/delete — dispatch by __type --

    vm.register_host_fn("vybe:collections", "collHas", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let type_str = o.properties.get("__type").map(|v| format!("{}", v)).unwrap_or_default();
            if type_str == "Map" {
                let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                if let Some(Value::Object(data)) = o.properties.get("__data") {
                    return Value::Bool(data.lock().unwrap().properties.contains_key(&key));
                }
            } else if type_str == "Set" {
                let value_str = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                if let Some(Value::Object(items)) = o.properties.get("__items") {
                    let items_ref = items.lock().unwrap();
                    if let ObjectKind::Array(ref elems) = items_ref.kind {
                        return Value::Bool(elems.iter().any(|e| format!("{}", e) == value_str));
                    }
                }
            }
        }
        Value::Bool(false)
    }));

    vm.register_host_fn("vybe:collections", "collDelete", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let type_str = {
                let o = obj.lock().unwrap();
                o.properties.get("__type").map(|v| format!("{}", v)).unwrap_or_default()
            };
            if type_str == "Map" {
                let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                let o = obj.lock().unwrap();
                if let Some(Value::Object(data)) = o.properties.get("__data") {
                    let existed = data.lock().unwrap().properties.remove(&key).is_some();
                    let size = data.lock().unwrap().properties.len() as f64;
                    drop(o);
                    obj.lock().unwrap().properties.insert("size".into(), Value::F64(size));
                    return Value::Bool(existed);
                }
            } else if type_str == "Set" {
                let value_str = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                let o = obj.lock().unwrap();
                if let Some(Value::Object(items)) = o.properties.get("__items") {
                    let mut items_ref = items.lock().unwrap();
                    if let ObjectKind::Array(ref mut elems) = items_ref.kind {
                        let before = elems.len();
                        elems.retain(|e| format!("{}", e) != value_str);
                        let removed = elems.len() < before;
                        let len = elems.len() as f64;
                        drop(items_ref);
                        drop(o);
                        obj.lock().unwrap().properties.insert("size".into(), Value::F64(len));
                        return Value::Bool(removed);
                    }
                }
            }
        }
        Value::Bool(false)
    }));
}
