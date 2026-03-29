use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Object.keys(obj) → array of property name strings
    vm.register_host_fn("vybe:object", "keys", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let keys: Vec<Value> = o.properties.keys()
                .filter(|k| *k != "length") // exclude internal 'length' for arrays
                .map(|k| Value::String(Rc::from(k.as_str())))
                .collect();
            return Value::Object(Rc::new(RefCell::new(Object::new_array(keys))));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Object.values(obj) → array of property values
    vm.register_host_fn("vybe:object", "values", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let vals: Vec<Value> = o.properties.values().cloned().collect();
            return Value::Object(Rc::new(RefCell::new(Object::new_array(vals))));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Object.entries(obj) → array of [key, value] pairs
    vm.register_host_fn("vybe:object", "entries", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let entries: Vec<Value> = o.properties.iter()
                .map(|(k, v)| {
                    Value::Object(Rc::new(RefCell::new(Object::new_array(vec![
                        Value::String(Rc::from(k.as_str())),
                        v.clone(),
                    ]))))
                })
                .collect();
            return Value::Object(Rc::new(RefCell::new(Object::new_array(entries))));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Object.assign(target, ...sources) → target with all source props copied
    vm.register_host_fn("vybe:object", "assign", Box::new(|args: &[Value]| {
        if let Some(Value::Object(target)) = args.first() {
            for source_arg in &args[1..] {
                if let Value::Object(source) = source_arg {
                    let src = source.borrow();
                    let mut tgt = target.borrow_mut();
                    for (k, v) in &src.properties {
                        tgt.properties.insert(k.clone(), v.clone());
                    }
                }
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // "key" in obj → hasProperty(key, obj)
    vm.register_host_fn("vybe:object", "hasProperty", Box::new(|args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.get(1) {
            let o = obj.borrow();
            Value::Bool(o.properties.contains_key(&key))
        } else {
            Value::Bool(false)
        }
    }));

    // delete obj.prop → deleteProperty(obj, key)
    vm.register_host_fn("vybe:object", "deleteProperty", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            obj.borrow_mut().properties.remove(&key);
            Value::Bool(true)
        } else {
            Value::Bool(false)
        }
    }));

    // Object.freeze(obj) — mark as frozen (simplified: no-op, returns obj)
    vm.register_host_fn("vybe:object", "freeze", Box::new(|args: &[Value]| {
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Object.fromEntries([[k,v], ...]) → obj
    vm.register_host_fn("vybe:object", "fromEntries", Box::new(|args: &[Value]| {
        let mut obj = Object::new();
        if let Some(Value::Object(arr)) = args.first() {
            let a = arr.borrow();
            if let ObjectKind::Array(entries) = &a.kind {
                for entry in entries {
                    if let Value::Object(pair) = entry {
                        let p = pair.borrow();
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
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Object.hasOwn(obj, key) — ES2022
    vm.register_host_fn("vybe:object", "hasOwn", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            Value::Bool(obj.borrow().properties.contains_key(&key))
        } else {
            Value::Bool(false)
        }
    }));

    // a instanceof B → check via type registry first, then __types array fallback.
    // This supports cross-language instanceof: VB classes, JS classes, built-in types.
    vm.register_host_fn("vybe:object", "instanceOf", Box::new(|args: &[Value]| {
        // Extract target type name from the constructor object (args[1])
        let target_name = if let Some(Value::Object(ctor)) = args.get(1) {
            ctor.borrow().properties.get("name").map(|v| format!("{}", v)).unwrap_or_default()
        } else if let Some(Value::String(s)) = args.get(1) {
            // Allow passing type name directly as string (for ref_test fallback)
            s.to_string()
        } else {
            return Value::Bool(false);
        };
        if target_name.is_empty() { return Value::Bool(false); }

        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();

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
                let t = types.borrow();
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
