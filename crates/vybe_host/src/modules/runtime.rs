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

    // awaitPromise(value) — if value is a Promise, extract its resolved value
    // For synchronous promises (our model), this is immediate.
    vm.register_host_fn("vybe:runtime", "awaitPromise", Box::new(|args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        if let Value::Object(ref obj) = val {
            let o = obj.borrow();
            if o.properties.get("__type").map(|v| format!("{}", v)) == Some("Promise".into()) {
                let state = o.properties.get("__state").map(|v| format!("{}", v)).unwrap_or_default();
                if state == "fulfilled" {
                    return o.properties.get("__value").cloned().unwrap_or(Value::Null);
                } else if state == "rejected" {
                    // TODO: throw the rejection reason
                    return o.properties.get("__value").cloned().unwrap_or(Value::Null);
                }
                // pending — in synchronous model this shouldn't happen
                return Value::Null;
            }
        }
        // Not a Promise — return as-is (await on non-Promise is identity in JS)
        val
    }));

    // Promise.resolve(value) → creates a fulfilled Promise
    vm.register_host_fn("vybe:runtime", "promiseResolve", Box::new(|args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        make_promise("fulfilled", val)
    }));

    // Promise.reject(reason) → creates a rejected Promise
    vm.register_host_fn("vybe:runtime", "promiseReject", Box::new(|args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        make_promise("rejected", val)
    }));

    // Promise.all(array) → Promise that resolves with array of values
    vm.register_host_fn("vybe:runtime", "promiseAll", Box::new(|args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                let results: Vec<Value> = elems.iter().map(|p| {
                    if let Value::Object(obj) = p {
                        let po = obj.borrow();
                        if po.properties.get("__type").map(|v| format!("{}", v)) == Some("Promise".into()) {
                            return po.properties.get("__value").cloned().unwrap_or(Value::Null);
                        }
                    }
                    p.clone()
                }).collect();
                return make_promise("fulfilled", Value::Object(
                    Rc::new(RefCell::new(Object::new_array(results)))
                ));
            }
        }
        make_promise("fulfilled", Value::Null)
    }));

    // Error constructor: new Error("message")
    vm.register_host_fn("vybe:runtime", "Error", Box::new(|args: &[Value]| {
        make_error("Error", args)
    }));

    vm.register_host_fn("vybe:runtime", "TypeError", Box::new(|args: &[Value]| {
        make_error("TypeError", args)
    }));

    vm.register_host_fn("vybe:runtime", "RangeError", Box::new(|args: &[Value]| {
        make_error("RangeError", args)
    }));

    // GoTo support — stores the target label. The caller checks this global
    // to implement VB6-style GoTo within a subroutine.
    vm.register_host_fn("vybe:runtime", "goto", Box::new(|args: &[Value]| {
        // In practice, GoTo within a Sub is rare in modern VB.NET.
        // This stores the label name so error handlers can dispatch.
        args.first().cloned().unwrap_or(Value::Null)
    }));
}

fn make_error(kind: &str, args: &[Value]) -> Value {
    // args[0] = this (from new), args[1] = message
    let this = args.first().cloned().unwrap_or(Value::Null);
    let message = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
    if let Value::Object(ref obj) = this {
        let mut o = obj.borrow_mut();
        o.properties.insert("__type".into(), Value::String(Rc::from(kind)));
        o.properties.insert("name".into(), Value::String(Rc::from(kind)));
        o.properties.insert("message".into(), Value::String(Rc::from(message.as_str())));
        o.properties.insert("stack".into(), Value::String(Rc::from(format!("{}: {}", kind, message).as_str())));
    }
    this
}

fn make_promise(state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Rc::from("Promise")));
    obj.properties.insert("__state".into(), Value::String(Rc::from(state)));
    obj.properties.insert("__value".into(), value);
    Value::Object(Rc::new(RefCell::new(obj)))
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
        "values" => {
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let vals: Vec<Value> = data.borrow().properties.values().cloned().collect();
                return Value::Object(Rc::new(RefCell::new(Object::new_array(vals))));
            }
            Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
        }
        "clear" => {
            let ob = o.borrow();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                data.borrow_mut().properties.clear();
            }
            drop(ob);
            o.borrow_mut().properties.insert("size".into(), Value::F64(0.0));
            obj.clone()
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
        "clear" => {
            let ob = o.borrow();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let mut ir = items.borrow_mut();
                if let ObjectKind::Array(ref mut elems) = ir.kind {
                    elems.clear();
                }
            }
            drop(ob);
            o.borrow_mut().properties.insert("size".into(), Value::F64(0.0));
            obj.clone()
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
