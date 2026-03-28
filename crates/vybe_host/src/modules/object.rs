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

    // Object.assign(target, source) → target with source props copied
    vm.register_host_fn("vybe:object", "assign", Box::new(|args: &[Value]| {
        if let (Some(Value::Object(target)), Some(Value::Object(source))) = (args.first(), args.get(1)) {
            let src = source.borrow();
            let mut tgt = target.borrow_mut();
            for (k, v) in &src.properties {
                tgt.properties.insert(k.clone(), v.clone());
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

    // a instanceof B → check if B.name is in a's type ancestry
    // Walk a's constructor chain (__super on instance → constructor, then __parent on constructors)
    vm.register_host_fn("vybe:object", "instanceOf", Box::new(|args: &[Value]| {
        let target_name = if let Some(Value::Object(ctor)) = args.get(1) {
            ctor.borrow().properties.get("name").map(|v| format!("{}", v)).unwrap_or_default()
        } else {
            return Value::Bool(false);
        };
        if target_name.is_empty() { return Value::Bool(false); }

        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            // Direct __type check
            if let Some(t) = o.properties.get("__type") {
                if format!("{}", t) == target_name { return Value::Bool(true); }
            }
            // Walk the __super chain on the instance (each __super is a parent constructor)
            let mut current = o.properties.get("__super").cloned();
            drop(o);
            for _ in 0..20 {
                let next = if let Some(Value::Object(ref sup)) = current {
                    let s = sup.borrow();
                    if let Some(n) = s.properties.get("name") {
                        if format!("{}", n) == target_name { return Value::Bool(true); }
                    }
                    s.properties.get("__parent").cloned()
                } else {
                    break;
                };
                current = next;
            }
        }
        Value::Bool(false)
    }));
}
