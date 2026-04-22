use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

/// Runtime dispatch — handles method calls that need type-aware routing.
/// Used when the compiler can't determine at compile time whether
/// obj.method() is a user method or a builtin (Map.set vs Class.set).
pub fn register(vm: &mut VM) {
    // callMethod(obj, methodName, ...args)
    // Routes to the right implementation based on obj's __type.
    vm.register_host_fn("vybe:runtime", "callMethod", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let obj = args.first().cloned().unwrap_or(Value::Null);
        let method = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let call_args = if args.len() > 2 { &args[2..] } else { &[] };

        if let Value::Object(ref o) = obj {
            let type_str = {
                let ob = o.lock().unwrap();
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

            // List methods (includes LINQ)
            if type_str == "List" {
                return dispatch_list(ctx, &obj, &method, call_args);
            }

            // .NET Dictionary methods
            if type_str == "Dictionary" {
                return dispatch_dictionary(&obj, &method, call_args);
            }
        }

        // Not a builtin collection — return Undefined (caller falls through to regular method call)
        Value::Undefined
    }));

    // awaitPromise(value) — if value is a Promise, extract its resolved value
    // For synchronous promises (our model), this is immediate.
    vm.register_host_fn("vybe:runtime", "awaitPromise", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        if let Value::Object(ref obj) = val {
            let o = obj.lock().unwrap();
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
    vm.register_host_fn("vybe:runtime", "promiseResolve", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        make_promise("fulfilled", val)
    }));

    // Promise.reject(reason) → creates a rejected Promise
    vm.register_host_fn("vybe:runtime", "promiseReject", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().cloned().unwrap_or(Value::Null);
        make_promise("rejected", val)
    }));

    // Promise.all(array) → Promise that resolves with array of values
    vm.register_host_fn("vybe:runtime", "promiseAll", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let results: Vec<Value> = elems.iter().map(|p| {
                    if let Value::Object(obj) = p {
                        let po = obj.lock().unwrap();
                        if po.properties.get("__type").map(|v| format!("{}", v)) == Some("Promise".into()) {
                            return po.properties.get("__value").cloned().unwrap_or(Value::Null);
                        }
                    }
                    p.clone()
                }).collect();
                return make_promise("fulfilled", Value::Object(
                    Arc::new(Mutex::new(Object::new_array(results)))
                ));
            }
        }
        make_promise("fulfilled", Value::Null)
    }));

    // Error constructor: new Error("message")
    vm.register_host_fn("vybe:runtime", "Error", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        make_error("Error", args)
    }));

    vm.register_host_fn("vybe:runtime", "TypeError", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        make_error("TypeError", args)
    }));

    vm.register_host_fn("vybe:runtime", "RangeError", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        make_error("RangeError", args)
    }));

    // GoTo support — stores the target label. The caller checks this global
    // to implement VB6-style GoTo within a subroutine.
    vm.register_host_fn("vybe:runtime", "goto", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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
        let mut o = obj.lock().unwrap();
        o.properties.insert("__type".into(), Value::String(Arc::from(kind)));
        o.properties.insert("__exception_type".into(), Value::String(Arc::from(kind)));
        o.properties.insert("name".into(), Value::String(Arc::from(kind)));
        o.properties.insert("message".into(), Value::String(Arc::from(message.as_str())));
        o.properties.insert("stack".into(), Value::String(Arc::from(format!("{}: {}", kind, message).as_str())));
    }
    this
}

fn make_promise(state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties.insert("__state".into(), Value::String(Arc::from(state)));
    obj.properties.insert("__value".into(), value);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn dispatch_list(ctx: &mut HostContext, obj: &Value, method: &str, args: &[Value]) -> Value {
    let o = match obj { Value::Object(o) => o, _ => return Value::Null };

    match method {
        "Add" | "add" => {
            let item = args.first().cloned().unwrap_or(Value::Null);
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                elems.push(item);
                let len = elems.len() as f64;
                ob.properties.insert("count".into(), Value::F64(len));
                ob.properties.insert("length".into(), Value::F64(len));
            }
            Value::Null // void return — handled
        }
        "Remove" | "remove" => {
            let item_str = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                let before = elems.len();
                elems.retain(|e| format!("{}", e) != item_str);
                let removed = elems.len() < before;
                let len = elems.len() as f64;
                ob.properties.insert("count".into(), Value::F64(len));
                ob.properties.insert("length".into(), Value::F64(len));
                return Value::Bool(removed);
            }
            Value::Bool(false)
        }
        "RemoveAt" | "removeat" => {
            let idx = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                if idx < elems.len() {
                    elems.remove(idx);
                    let len = elems.len() as f64;
                    ob.properties.insert("count".into(), Value::F64(len));
                    ob.properties.insert("length".into(), Value::F64(len));
                }
            }
            Value::Null
        }
        "Clear" | "clear" => {
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                elems.clear();
                ob.properties.insert("count".into(), Value::F64(0.0));
                ob.properties.insert("length".into(), Value::F64(0.0));
            }
            Value::Null
        }
        "Contains" | "contains" => {
            let search = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search));
            }
            Value::Bool(false)
        }
        "IndexOf" | "indexof" => {
            let search = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                for (i, e) in elems.iter().enumerate() {
                    if format!("{}", e) == search { return Value::F64(i as f64); }
                }
            }
            Value::F64(-1.0)
        }
        "Insert" | "insert" => {
            let idx = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            let item = args.get(1).cloned().unwrap_or(Value::Null);
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                if idx <= elems.len() { elems.insert(idx, item); }
                let len = elems.len() as f64;
                ob.properties.insert("count".into(), Value::F64(len));
                ob.properties.insert("length".into(), Value::F64(len));
            }
            Value::Null
        }
        "Item" | "item" | "get_Item" => {
            let idx = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                return elems.get(idx).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }
        "Count" | "count" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                return Value::F64(elems.len() as f64);
            }
            Value::F64(0.0)
        }
        "Reverse" | "reverse" => {
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                elems.reverse();
            }
            obj.clone()
        }
        "Sort" | "sort" => {
            let mut ob = o.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = ob.kind {
                elems.sort_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal));
            }
            obj.clone()
        }
        "ToArray" | "toarray" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
        // LINQ methods — work on any List/Array with VM callback support
        "Where" | "where" => {
            // Where(predicate) — filter using VM callback
            if args.is_empty() { return Value::Null; }
            let predicate = &args[0];
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let elems_clone = elems.clone();
                drop(ob);
                let mut filtered = Vec::new();
                for elem in &elems_clone {
                    let result = ctx.invoke(predicate, &[elem.clone()]);
                    if result.as_bool() {
                        filtered.push(elem.clone());
                    }
                }
                let mut result_obj = Object::new_array(filtered);
                result_obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(result_obj)));
            }
            drop(ob);
            Value::Null
        }
        "Select" | "select" => {
            // Select(mapper) — map using VM callback
            if args.is_empty() { return Value::Null; }
            let mapper = &args[0];
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let elems_clone = elems.clone();
                drop(ob);
                let mut mapped = Vec::new();
                for elem in &elems_clone {
                    let result = ctx.invoke(mapper, &[elem.clone()]);
                    mapped.push(result);
                }
                let mut result_obj = Object::new_array(mapped);
                result_obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(result_obj)));
            }
            drop(ob);
            Value::Null
        }
        "First" | "first" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                return elems.first().cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }
        "Last" | "last" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                return elems.last().cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }
        "Any" | "any" => {
            if !args.is_empty() {
                // Any(predicate) — check if any element matches
                let predicate = &args[0];
                let ob = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ob.kind {
                    let elems_clone = elems.clone();
                    drop(ob);
                    for elem in &elems_clone {
                        let result = ctx.invoke(predicate, &[elem.clone()]);
                        if result.as_bool() { return Value::Bool(true); }
                    }
                    return Value::Bool(false);
                }
                drop(ob);
                Value::Bool(false)
            } else {
                // Any() — check if non-empty
                let ob = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ob.kind {
                    return Value::Bool(!elems.is_empty());
                }
                Value::Bool(false)
            }
        }
        "All" | "all" => {
            if !args.is_empty() {
                let predicate = &args[0];
                let ob = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ob.kind {
                    let elems_clone = elems.clone();
                    drop(ob);
                    for elem in &elems_clone {
                        let result = ctx.invoke(predicate, &[elem.clone()]);
                        if !result.as_bool() { return Value::Bool(false); }
                    }
                    return Value::Bool(true);
                }
                drop(ob);
            }
            Value::Bool(true)
        }
        "ForEach" | "forEach" | "foreach" => {
            if !args.is_empty() {
                let callback = &args[0];
                let ob = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ob.kind {
                    let elems_clone = elems.clone();
                    drop(ob);
                    for elem in &elems_clone {
                        ctx.invoke(callback, &[elem.clone()]);
                    }
                    return Value::Null;
                }
                drop(ob);
            }
            Value::Null
        }
        "Aggregate" | "aggregate" | "Reduce" | "reduce" => {
            if !args.is_empty() {
                let reducer = &args[0];
                let ob = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ob.kind {
                    if elems.is_empty() { drop(ob); return Value::Null; }
                    let elems_clone = elems.clone();
                    drop(ob);
                    let mut acc = elems_clone[0].clone();
                    for elem in &elems_clone[1..] {
                        acc = ctx.invoke(reducer, &[acc, elem.clone()]);
                    }
                    return acc;
                }
                drop(ob);
            }
            Value::Null
        }
        "OrderBy" | "orderBy" | "orderby" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let mut sorted = elems.clone();
                sorted.sort_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal));
                let mut result_obj = Object::new_array(sorted);
                result_obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(result_obj)));
            }
            Value::Null
        }
        "Sum" | "sum" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let sum: f64 = elems.iter().map(|e| e.as_f64()).sum();
                return Value::F64(sum);
            }
            Value::F64(0.0)
        }
        "Min" | "min" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                if let Some(min) = elems.iter().map(|e| e.as_f64()).reduce(f64::min) {
                    return Value::F64(min);
                }
            }
            Value::Null
        }
        "Max" | "max" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                if let Some(max) = elems.iter().map(|e| e.as_f64()).reduce(f64::max) {
                    return Value::F64(max);
                }
            }
            Value::Null
        }
        "Average" | "average" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                if !elems.is_empty() {
                    let sum: f64 = elems.iter().map(|e| e.as_f64()).sum();
                    return Value::F64(sum / elems.len() as f64);
                }
            }
            Value::Null
        }
        "Distinct" | "distinct" => {
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let mut seen = std::collections::HashSet::new();
                let mut result = Vec::new();
                for e in elems {
                    let s = format!("{}", e);
                    if seen.insert(s) { result.push(e.clone()); }
                }
                let mut new_obj = Object::new_array(result);
                new_obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(new_obj)));
            }
            obj.clone()
        }
        "Take" | "take" => {
            let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let taken: Vec<Value> = elems.iter().take(n).cloned().collect();
                let mut new_obj = Object::new_array(taken);
                new_obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(new_obj)));
            }
            obj.clone()
        }
        "Skip" | "skip" => {
            let n = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
            let ob = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = ob.kind {
                let skipped: Vec<Value> = elems.iter().skip(n).cloned().collect();
                let mut new_obj = Object::new_array(skipped);
                new_obj.properties.insert("__type".into(), Value::String(Arc::from("List")));
                return Value::Object(Arc::new(Mutex::new(new_obj)));
            }
            obj.clone()
        }
        _ => Value::Undefined, // not handled — compiler falls through to struct_get
    }
}

fn dispatch_map(obj: &Value, method: &str, args: &[Value]) -> Value {
    let o = match obj { Value::Object(o) => o, _ => return Value::Null };

    match method {
        "set" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(1).cloned().unwrap_or(Value::Null);
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                data.lock().unwrap().properties.insert(key, value);
                let size = data.lock().unwrap().properties.len() as f64;
                drop(ob);
                o.lock().unwrap().properties.insert("size".into(), Value::F64(size));
            }
            obj.clone()
        }
        "get" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }
        "has" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return Value::Bool(data.lock().unwrap().properties.contains_key(&key));
            }
            Value::Bool(false)
        }
        "delete" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let existed = data.lock().unwrap().properties.remove(&key).is_some();
                let size = data.lock().unwrap().properties.len() as f64;
                drop(ob);
                o.lock().unwrap().properties.insert("size".into(), Value::F64(size));
                return Value::Bool(existed);
            }
            Value::Bool(false)
        }
        "keys" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let keys: Vec<Value> = data.lock().unwrap().properties.keys()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
        "values" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let vals: Vec<Value> = data.lock().unwrap().properties.values().cloned().collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
        "clear" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                data.lock().unwrap().properties.clear();
            }
            drop(ob);
            o.lock().unwrap().properties.insert("size".into(), Value::F64(0.0));
            obj.clone()
        }
        _ => Value::Null,
    }
}

fn dispatch_dictionary(obj: &Value, method: &str, args: &[Value]) -> Value {
    let o = match obj { Value::Object(o) => o, _ => return Value::Null };

    match method {
        "Add" | "add" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let value = args.get(1).cloned().unwrap_or(Value::Null);
            let data = {
                let ob = o.lock().unwrap();
                match ob.properties.get("__data") {
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
                let mut ob = o.lock().unwrap();
                if !key.starts_with("__") {
                    ob.properties.insert(key, value);
                }
                ob.properties.insert("count".into(), Value::F64(count));
                ob.properties.insert("length".into(), Value::F64(count));
            }
            Value::Null
        }
        "Item" | "item" | "get_Item" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            ob.properties.get(&key).cloned().unwrap_or(Value::Null)
        }
        "ContainsKey" | "containskey" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return Value::Bool(data.lock().unwrap().properties.contains_key(&key));
            }
            Value::Bool(ob.properties.contains_key(&key))
        }
        "Remove" | "remove" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let data = {
                let ob = o.lock().unwrap();
                match ob.properties.get("__data") {
                    Some(Value::Object(data)) => Some(data.clone()),
                    _ => None,
                }
            };
            if let Some(data) = data {
                let (removed, count) = {
                    let mut data_obj = data.lock().unwrap();
                    let removed = data_obj.properties.remove(&key).is_some();
                    (removed, data_obj.properties.len() as f64)
                };
                let mut ob = o.lock().unwrap();
                ob.properties.remove(&key);
                ob.properties.insert("count".into(), Value::F64(count));
                ob.properties.insert("length".into(), Value::F64(count));
                return Value::Bool(removed);
            }
            Value::Bool(false)
        }
        "Keys" | "keys" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let keys: Vec<Value> = data.lock().unwrap().properties.keys()
                    .map(|k| Value::String(Arc::from(k.as_str())))
                    .collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
        "Values" | "values" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                let vals: Vec<Value> = data.lock().unwrap().properties.values().cloned().collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(vals))));
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
        "Clear" | "clear" => {
            let data = {
                let ob = o.lock().unwrap();
                match ob.properties.get("__data") {
                    Some(Value::Object(data)) => Some(data.clone()),
                    _ => None,
                }
            };
            if let Some(data) = data {
                data.lock().unwrap().properties.clear();
                let mut ob = o.lock().unwrap();
                ob.properties.retain(|k, _| k == "__type" || k == "__data" || k == "count" || k == "length");
                ob.properties.insert("count".into(), Value::F64(0.0));
                ob.properties.insert("length".into(), Value::F64(0.0));
            }
            Value::Null
        }
        "Count" | "count" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return Value::F64(data.lock().unwrap().properties.len() as f64);
            }
            ob.properties.get("count").cloned().unwrap_or(Value::F64(0.0))
        }
        "TryGetValue" | "trygetvalue" => {
            let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(data)) = ob.properties.get("__data") {
                return data.lock().unwrap().properties.get(&key).cloned().unwrap_or(Value::Null);
            }
            ob.properties.get(&key).cloned().unwrap_or(Value::Null)
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
            let ob = o.lock().unwrap();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let mut ir = items.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = ir.kind {
                    if !elems.iter().any(|e| format!("{}", e) == value_str) {
                        elems.push(value);
                        let len = elems.len() as f64;
                        drop(ir);
                        drop(ob);
                        o.lock().unwrap().properties.insert("size".into(), Value::F64(len));
                    }
                }
            }
            obj.clone()
        }
        "has" => {
            let value_str = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let ir = items.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ir.kind {
                    return Value::Bool(elems.iter().any(|e| format!("{}", e) == value_str));
                }
            }
            Value::Bool(false)
        }
        "delete" => {
            let value_str = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let ob = o.lock().unwrap();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let mut ir = items.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = ir.kind {
                    let before = elems.len();
                    elems.retain(|e| format!("{}", e) != value_str);
                    let removed = elems.len() < before;
                    let len = elems.len() as f64;
                    drop(ir);
                    drop(ob);
                    o.lock().unwrap().properties.insert("size".into(), Value::F64(len));
                    return Value::Bool(removed);
                }
            }
            Value::Bool(false)
        }
        "clear" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let mut ir = items.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = ir.kind {
                    elems.clear();
                }
            }
            drop(ob);
            o.lock().unwrap().properties.insert("size".into(), Value::F64(0.0));
            obj.clone()
        }
        "values" => {
            let ob = o.lock().unwrap();
            if let Some(Value::Object(items)) = ob.properties.get("__items") {
                let ir = items.lock().unwrap();
                if let ObjectKind::Array(ref elems) = ir.kind {
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
                }
            }
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
        _ => Value::Null,
    }
}
