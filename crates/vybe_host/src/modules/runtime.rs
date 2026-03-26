use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};

/// Runtime dispatch — handles method calls that need type-aware routing.
/// Used when the compiler can't determine at compile time whether
/// obj.method() is a user method or a builtin (Map.set vs Class.set).
pub fn register(vm: &mut VM) {
    // callMethod(obj, methodName, ...args)
    // Routes to the right implementation based on obj's __type.
    vm.register_host_fn("vybe:runtime", "callMethod", Box::new(|args: &[Value]| {
        let obj = args.first().cloned().unwrap_or(Value::Null);
        let method = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let call_args = if args.len() > 2 { &args[2..] } else { &[] };

        if let Value::Object(ref o) = obj {
            let type_str = {
                let ob = o.borrow();
                ob.properties.get("__type").map(|v| format!("{}", v)).unwrap_or_default()
            };

            // Map methods
            if type_str == "Map" {
                return dispatch_map(&obj, &method, call_args);
            }

            // Set methods
            if type_str == "Set" {
                return dispatch_set(&obj, &method, call_args);
            }
        }

        // Not a builtin collection — return Null (caller falls through to regular method call)
        Value::Null
    }));
}

fn dispatch_map(obj: &Value, method: &str, args: &[Value]) -> Value {
    let o = match obj { Value::Object(o) => o, _ => return Value::Null };

    match method {
        "set" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(1).cloned().unwrap_or(Value::Null);
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                data.borrow_mut().properties.insert(key, value);
                let size = data.borrow().properties.len() as f64;
                drop(ob);
                o.borrow_mut().properties.insert("size".into(), Value::F64(size));
            }
            obj.clone()
        }
        "get" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return data.borrow().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }
        "has" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return Value::Bool(data.borrow().properties.contains_key(&key));
            }
            Value::Bool(false)
        }
        "delete" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let existed = data.borrow_mut().properties.remove(&key).is_some();
                let size = data.borrow().properties.len() as f64;
                drop(ob);
                o.borrow_mut().properties.insert("size".into(), Value::F64(size));
                return Value::Bool(existed);
            }
            Value::Bool(false)
        }
        "keys" => {
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let keys: Vec<Value> = data.borrow().properties.keys()
                    .map(|k| Value::String(Rc::from(k.as_str())))
                    .collect();
                return Value::Object(Rc::new(RefCell::new(Object::new_array(keys))));
            }
            Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
        }
        _ => Value::Null,
    }
}

fn dispatch_set(obj: &Value, method: &str, args: &[Value]) -> Value {
    let o = match obj { Value::Object(o) => o, _ => return Value::Null };

    match method {
        "add" => {
            let value = args.first().cloned().unwrap_or(Value::Null);
            let value_str = format!("{}", value);
            let ob = o.borrow();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let mut ir = items.borrow_mut();
                if let ObjectKind::Array(ref mut elems) = ir.kind {
                    if !elems.iter().any(|e| format!("{}", e) == value_str) {
                        elems.push(value);
                        let len = elems.len() as f64;
                        drop(ir);
                        drop(ob);
                        o.borrow_mut().properties.insert("size".into(), Value::F64(len));
                    }
                }
            }
            obj.clone()
        }
        "has" => {
            let value_str = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.borrow();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let ir = items.borrow();
                if let ObjectKind::Array(ref elems) = ir.kind {
                    return Value::Bool(elems.iter().any(|e| format!("{}", e) == value_str));
                }
            }
            Value::Bool(false)
        }
        "delete" => {
            let value_str = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.borrow();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let mut ir = items.borrow_mut();
                if let ObjectKind::Array(ref mut elems) = ir.kind {
                    let before = elems.len();
                    elems.retain(|e| format!("{}", e) != value_str);
                    let removed = elems.len() < before;
                    let len = elems.len() as f64;
                    drop(ir);
                    drop(ob);
                    o.borrow_mut().properties.insert("size".into(), Value::F64(len));
                    return Value::Bool(removed);
                }
            }
            Value::Bool(false)
        }
        "values" => {
            let ob = o.borrow();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let ir = items.borrow();
                if let ObjectKind::Array(ref elems) = ir.kind {
                    return Value::Object(Rc::new(RefCell::new(Object::new_array(elems.clone()))));
                }
            }
            Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
        }
        _ => Value::Null,
    }
}
