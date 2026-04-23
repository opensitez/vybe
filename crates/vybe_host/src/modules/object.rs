use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // `vybe:object.keys/values/entries` — POLYMORPHIC iteration primitives.
    //
    // The same three host fns serve every language's iteration needs:
    // - JS `Object.keys/values/entries`
    // - Python `dict.keys/values/items`
    // - PHP `array_keys/array_values/array_map`-shaped iteration, `foreach`
    // - Ruby `Hash#keys/values/to_a`
    // - C# `Dictionary<K,V>.Keys/Values/KeyValuePairs`
    // - Dart `Map.keys/values/entries`
    //
    // All dispatch on the value's actual type:
    //   - `ObjectKind::Array(v)`  → integer-indexed values
    //   - `ObjectKind::Map(m)`    → canonical associative (IndexMap) — PHP
    //                                assoc, Python dict, Ruby Hash, JS object
    //                                literal with string keys
    //   - `ObjectKind::Ordinary`  → property bag (JS plain object, class
    //                                instances)
    //   - other kinds (TypedArray, Set, …) → fall back to empty Array
    //
    // This is the single polymorphic dispatch that the compiler_common
    // iteration emitters depend on — one impl, every language benefits.

    vm.register_host_fn("vybe:object", "keys", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(v) => {
                    // Integer indices as string keys (matches JS
                    // `Object.keys([a,b,c])` = ["0","1","2"]).
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
            // Ordinary fallback — honors __keys marker for insertion order.
            if let Some(Value::Object(keys_arr)) = o.properties.get("__keys") {
                let ka = keys_arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ka.kind {
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
                }
            }
            let keys: Vec<Value> = o.properties.keys()
                .filter(|k| *k != "length" && !k.starts_with("__"))
                .map(|k| Value::String(Arc::from(k.as_str())))
                .collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    vm.register_host_fn("vybe:object", "values", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
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
            if let Some(Value::Object(keys_arr)) = o.properties.get("__keys") {
                let ka = keys_arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ka.kind {
                    let vals: Vec<Value> = elems.iter()
                        .filter_map(|k| if let Value::String(s) = k { o.properties.get(s.as_ref()).cloned() } else { None })
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
                }
            }
            let vals: Vec<Value> = o.properties.iter()
                .filter(|(k, _)| !k.starts_with("__"))
                .map(|(_, v)| v.clone())
                .collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    vm.register_host_fn("vybe:object", "entries", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(v) => {
                    // Integer index + value pairs.
                    let entries: Vec<Value> = v.iter().enumerate()
                        .map(|(i, val)| {
                            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                                Value::I32(i as i32),
                                val.clone(),
                            ]))))
                        })
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                }
                ObjectKind::Map(m) => {
                    let entries: Vec<Value> = m.iter()
                        .map(|(k, v)| {
                            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                                k.clone(),
                                v.clone(),
                            ]))))
                        })
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                }
                _ => {}
            }
            if let Some(Value::Object(keys_arr)) = o.properties.get("__keys") {
                let ka = keys_arr.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ka.kind {
                    let entries: Vec<Value> = elems.iter()
                        .filter_map(|k| {
                            if let Value::String(s) = k {
                                o.properties.get(s.as_ref()).map(|v| {
                                    Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                                        Value::String(s.clone()),
                                        v.clone(),
                                    ]))))
                                })
                            } else { None }
                        })
                        .collect();
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
                }
            }
            let entries: Vec<Value> = o.properties.iter()
                .filter(|(k, _)| !k.starts_with("__"))
                .map(|(k, v)| {
                    Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                        Value::String(Arc::from(k.as_str())),
                        v.clone(),
                    ]))))
                })
                .collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(entries))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // Object.assign(target, ...sources) → target with all source props copied
    vm.register_host_fn("vybe:object", "assign", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(target)) = args.first() {
            for source_arg in &args[1..] {
                if let Value::Object(source) = source_arg {
                    let src = source.lock().unwrap();
                    let mut tgt = target.lock().unwrap();
                    for (k, v) in &src.properties {
                        tgt.properties.insert(k.clone(), v.clone());
                    }
                }
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // "key" in obj → hasProperty(key, obj)
    vm.register_host_fn("vybe:object", "hasProperty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.get(1) {
            let o = obj.lock().unwrap();
            Value::Bool(o.properties.contains_key(&key))
        } else {
            Value::Bool(false)
        }
    }));

    // delete obj.prop → deleteProperty(obj, key)
    vm.register_host_fn("vybe:object", "deleteProperty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            obj.lock().unwrap().properties.remove(&key);
            Value::Bool(true)
        } else {
            Value::Bool(false)
        }
    }));

    // Object.freeze(obj) — mark as frozen (simplified: no-op, returns obj)
    vm.register_host_fn("vybe:object", "freeze", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Object.create(proto, [props]) → new object inheriting from proto
    // Simplified: creates an empty object that copies proto's properties
    vm.register_host_fn("vybe:object", "create", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let mut obj = Object::new();
        if let Some(Value::Object(proto)) = args.first() {
            let p = proto.lock().unwrap();
            // Copy proto's properties as inherited
            for (k, v) in &p.properties {
                obj.properties.insert(k.clone(), v.clone());
            }
        }
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Object.seal(obj) → no-op (return same object)
    vm.register_host_fn("vybe:object", "seal", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Object.isFrozen(obj) → false (we don't track freeze state)
    vm.register_host_fn("vybe:object", "isFrozen", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Bool(false)
    }));

    // Object.isSealed(obj)
    vm.register_host_fn("vybe:object", "isSealed", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Bool(false)
    }));

    // Object.is(a, b) — like === but treats NaN==NaN and -0!==+0
    vm.register_host_fn("vybe:object", "is", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().cloned().unwrap_or(Value::Null);
        let b = args.get(1).cloned().unwrap_or(Value::Null);
        Value::Bool(format!("{:?}", a) == format!("{:?}", b))
    }));

    // Object.getPrototypeOf(obj) → null (we don't track prototypes)
    vm.register_host_fn("vybe:object", "getPrototypeOf", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Null
    }));

    // Object.getOwnPropertyNames(obj) → array of string keys (own properties only)
    vm.register_host_fn("vybe:object", "getOwnPropertyNames", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let mut names: Vec<Value> = o.properties.keys()
                .filter(|k| !k.starts_with("__"))
                .map(|k| Value::String(Arc::from(k.as_str())))
                .collect();
            names.sort_by(|a, b| format!("{}", a).cmp(&format!("{}", b)));
            let mut arr = Object::new();
            arr.kind = ObjectKind::Array(names);
            Value::Object(Arc::new(Mutex::new(arr)))
        } else {
            let mut arr = Object::new();
            arr.kind = ObjectKind::Array(vec![]);
            Value::Object(Arc::new(Mutex::new(arr)))
        }
    }));

    // Object.defineProperty(obj, key, descriptor) — simplified: set property to descriptor.value
    vm.register_host_fn("vybe:object", "defineProperty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let (Some(Value::Object(obj)), Some(key), Some(Value::Object(desc))) = (args.first(), args.get(1), args.get(2)) {
            let key_str = format!("{}", key);
            let d = desc.lock().unwrap();
            if let Some(val) = d.properties.get("value") {
                let mut o = obj.lock().unwrap();
                o.properties.insert(key_str, val.clone());
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Object.fromEntries([[k,v], ...]) → obj. Also accepts Map.
    vm.register_host_fn("vybe:object", "fromEntries", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let mut obj = Object::new();
        if let Some(Value::Object(arr)) = args.first() {
            let a = arr.lock().unwrap();
            // Map: copy from __data
            let type_name = a.properties.get("__type")
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            if type_name == "Map" {
                if let Some(Value::Object(data)) = a.properties.get("__data") {
                    let d = data.lock().unwrap();
                    for (k, v) in &d.properties {
                        obj.properties.insert(k.clone(), v.clone());
                    }
                }
            } else if let ObjectKind::Array(entries) = &a.kind {
                // Array of [k, v] pairs
                for entry in entries {
                    if let Value::Object(pair) = entry {
                        let p = pair.lock().unwrap();
                        if let ObjectKind::Array(kv) = &p.kind {
                            if kv.len() >= 2 {
                                let k = format!("{}", kv[0]);
                                obj.properties.insert(k, kv[1].clone());
                            }
                        }
                    }
                }
            }
        }
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Object.hasOwn(obj, key) — ES2022
    vm.register_host_fn("vybe:object", "hasOwn", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            Value::Bool(obj.lock().unwrap().properties.contains_key(&key))
        } else {
            Value::Bool(false)
        }
    }));

    // a instanceof B → check via type registry first, then __types array fallback.
    // This supports cross-language instanceof: VB classes, JS classes, built-in types.
    vm.register_host_fn("vybe:object", "instanceOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // Extract target type name from the constructor object (args[1])
        let target_name = if let Some(Value::Object(ctor)) = args.get(1) {
            let ob = ctor.lock().unwrap();
            // Try properties["name"] first, then Function.name
            ob.properties.get("name").map(|v| format!("{}", v))
                .or_else(|| {
                    if let ObjectKind::Function(ref f) = ob.kind {
                        f.name.clone()
                    } else { None }
                })
                .unwrap_or_default()
        } else if let Some(Value::String(s)) = args.get(1) {
            // Allow passing type name directly as string (for ref_test fallback)
            s.to_string()
        } else {
            return Value::Bool(false);
        };
        if target_name.is_empty() { return Value::Bool(false); }

        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();

            // 1. Try type_id-based check via type registry (fast path)
            //    This uses the same logic as ref_test/test_type in the VM.
            //    Objects with type_id > 0 have been registered in the type system.
            //    For type_id == 0 we still check __type string against the registry.
            let obj_type_name = o.properties.get("__type")
                .map(|v| format!("{}", v))
                .or_else(|| o.properties.get("__control_type")
                    .map(|v| format!("{}", v)))
                .unwrap_or_default();

            // Direct name match (case-insensitive)
            if obj_type_name.eq_ignore_ascii_case(&target_name) {
                return Value::Bool(true);
            }

            // 2. Check __types array (JS class inheritance chain)
            if let Some(Value::Object(types)) = o.properties.get("__types") {
                let t = types.lock().unwrap();
                if let ObjectKind::Array(ref elems) = t.kind {
                    if elems.iter().any(|e| format!("{}", e) == target_name) {
                        return Value::Bool(true);
                    }
                }
            }

            // 3. Fallback: check __type directly (legacy)
            if let Some(t) = o.properties.get("__type") {
                if format!("{}", t) == target_name {
                    return Value::Bool(true);
                }
            }
        }
        Value::Bool(false)
    }));
}
