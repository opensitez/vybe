use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // push, pop, shift, length, join, reverse, concat — removed (now direct VM opcodes)

    // redim(array, newSize, preserve) → resized array
    vm.register_host_fn("vybe:array", "redim", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let new_size = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let preserve = matches!(args.get(2), Some(Value::Bool(true)));
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                if preserve {
                    elems.resize(new_size, Value::Null);
                } else {
                    *elems = vec![Value::Null; new_size];
                }
            }
            drop(o);
            return args[0].clone();
        }
        // Not an array — create new one
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::Null; new_size]))))
    }));

    vm.register_host_fn("vybe:array", "indexOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let search = args.get(1).cloned().unwrap_or(Value::Null);
                for (i, elem) in elems.iter().enumerate() {
                    if elem.eq(&search) { return Value::F64(i as f64); }
                }
            }
        }
        Value::F64(-1.0)
    }));
    // reverse, concat — removed (now array_reverse, array_concat opcodes)

    vm.register_host_fn("vybe:array", "slice", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let len = elems.len() as i64;
                    let start = norm(args, 1, 0, len);
                    let end = norm(args, 2, len, len);
                    let sliced = if start < end { elems[start..end].to_vec() } else { vec![] };
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(sliced))));
                }
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
            }
            Some(Value::String(s)) => {
                let len = s.len() as i64;
                let start = norm(args, 1, 0, len);
                let end = norm(args, 2, len, len);
                if start < end { Value::String(Arc::from(&s[start..end])) }
                else { Value::String(Arc::from("")) }
            }
            _ => Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
    }));
    // sliceStep(arr, start, end, step) — slice with step, handles negative step for reverse
    vm.register_host_fn("vybe:array", "sliceStep", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let len = elems.len() as i64;
                    let step = args.get(3).map(|v| v.as_f64() as i64).unwrap_or(1);
                    if step == 0 { return Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))); }
                    let (default_start, default_end) = if step > 0 { (0i64, len) } else { (len - 1, -len - 1) };
                    let start = args.get(1).and_then(|v| if matches!(v, Value::Null) { None } else { Some(v.as_f64() as i64) }).unwrap_or(default_start);
                    let end = args.get(2).and_then(|v| if matches!(v, Value::Null) { None } else { Some(v.as_f64() as i64) }).unwrap_or(default_end);
                    let s = if start < 0 { (len + start).max(0) } else { start.min(len) };
                    let e = if end < 0 { (len + end).max(-1) } else { end.min(len) };
                    let mut result = Vec::new();
                    if step > 0 {
                        let mut i = s;
                        while i < e { result.push(elems[i as usize].clone()); i += step; }
                    } else {
                        let mut i = s;
                        while i > e { result.push(elems[i as usize].clone()); i += step; }
                    }
                    return Value::Object(Arc::new(Mutex::new(Object::new_array(result))));
                }
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
            }
            Some(Value::String(s)) => {
                let chars: Vec<char> = s.chars().collect();
                let len = chars.len() as i64;
                let step = args.get(3).map(|v| v.as_f64() as i64).unwrap_or(1);
                if step == 0 { return Value::String(Arc::from("")); }
                let (default_start, default_end) = if step > 0 { (0i64, len) } else { (len - 1, -len - 1) };
                let start = args.get(1).and_then(|v| if matches!(v, Value::Null) { None } else { Some(v.as_f64() as i64) }).unwrap_or(default_start);
                let end = args.get(2).and_then(|v| if matches!(v, Value::Null) { None } else { Some(v.as_f64() as i64) }).unwrap_or(default_end);
                let sv = if start < 0 { (len + start).max(0) } else { start.min(len) };
                let ev = if end < 0 { (len + end).max(-1) } else { end.min(len) };
                let mut result = String::new();
                if step > 0 {
                    let mut i = sv;
                    while i < ev { result.push(chars[i as usize]); i += step; }
                } else {
                    let mut i = sv;
                    while i > ev { result.push(chars[i as usize]); i += step; }
                }
                Value::String(Arc::from(result.as_str()))
            }
            _ => Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        }
    }));

    // dynMul(a, b) — dynamic multiply: str*int → repeat, int*int → multiply
    vm.register_host_fn("vybe:math", "dynMul", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().cloned().unwrap_or(Value::Null);
        let b = args.get(1).cloned().unwrap_or(Value::Null);
        match (&a, &b) {
            (Value::String(s), _) => {
                let count = b.as_f64() as usize;
                Value::String(Arc::from(s.repeat(count).as_str()))
            }
            (_, Value::String(s)) => {
                let count = a.as_f64() as usize;
                Value::String(Arc::from(s.repeat(count).as_str()))
            }
            (Value::I32(x), Value::I32(y)) => Value::I32(x.wrapping_mul(*y)),
            _ => Value::F64(a.as_f64() * b.as_f64()),
        }
    }));

    // setAt(arr, index, value) — set element at index
    vm.register_host_fn("vybe:array", "setAt", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let value = args.get(2).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                if idx < elems.len() {
                    elems[idx] = value;
                }
            }
        }
        Value::Null
    }));

    // fill(arr, value, start?, end?) → arr
    vm.register_host_fn("vybe:array", "fill", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let value = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                let len = elems.len();
                let start = args.get(2).map(|v| v.as_f64() as usize).unwrap_or(0).min(len);
                let end = args.get(3).map(|v| v.as_f64() as usize).unwrap_or(len).min(len);
                for i in start..end { elems[i] = value.clone(); }
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // flat(arr) → flattened array (one level)
    vm.register_host_fn("vybe:array", "flat", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let mut result = Vec::new();
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                for elem in elems {
                    if let Value::Object(inner) = elem {
                        let inner_obj = inner.lock().unwrap();
                        if let ObjectKind::Array(ref inner_elems) = inner_obj.kind {
                            result.extend(inner_elems.clone());
                            continue;
                        }
                    }
                    result.push(elem.clone());
                }
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(result))))
    }));

    // includes(arr, value) → bool
    vm.register_host_fn("vybe:array", "includes", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).cloned().unwrap_or(Value::Null);
            let search_str = format!("{}", search);
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search_str));
            }
        }
        Value::Bool(false)
    }));

    // Array.isArray(value)
    vm.register_host_fn("vybe:array", "isArray", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            return Value::Bool(matches!(o.kind, ObjectKind::Array(_)));
        }
        Value::Bool(false)
    }));

    // Array.from(arrayLike) — creates a new array from an iterable/array-like
    vm.register_host_fn("vybe:array", "from", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // --- VB-compatible array functions ---

    // ubound(arr) → last valid index (length - 1)
    vm.register_host_fn("vybe:array", "ubound", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::F64(elems.len() as f64 - 1.0);
            }
        }
        Value::F64(-1.0)
    }));

    // lbound(arr) → always 0
    vm.register_host_fn("vybe:array", "lbound", Box::new(|_ctx, _| Value::F64(0.0)));

    // unshift(arr, value) → prepend element, return new length
    vm.register_host_fn("vybe:array", "unshift", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                for i in (1..args.len()).rev() {
                    elems.insert(0, args[i].clone());
                }
                let len = elems.len() as f64;
                drop(o);
                obj.lock().unwrap().properties.insert("length".into(), Value::F64(len));
                return Value::F64(len);
            }
        }
        Value::Null
    }));

    // splice(arr, start, deleteCount, ...items) → removed elements
    vm.register_host_fn("vybe:array", "splice", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let start = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let delete_count = args.get(2).map(|v| v.as_f64() as usize).unwrap_or(0);
            let insert_items: Vec<Value> = args.iter().skip(3).cloned().collect();
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                let start = start.min(elems.len());
                let end = (start + delete_count).min(elems.len());
                let removed: Vec<Value> = elems.drain(start..end).collect();
                for (i, item) in insert_items.into_iter().enumerate() {
                    elems.insert(start + i, item);
                }
                let len = elems.len() as f64;
                drop(o);
                obj.lock().unwrap().properties.insert("length".into(), Value::F64(len));
                return Value::Object(Arc::new(Mutex::new(Object::new_array(removed))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // ── Python builtins ──────────────────────────────────────────

    // range(stop) or range(start, stop) or range(start, stop, step)
    vm.register_host_fn("vybe:array", "range", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (start, stop, step) = match args.len() {
            1 => (0i64, args[0].as_f64() as i64, 1i64),
            2 => (args[0].as_f64() as i64, args[1].as_f64() as i64, 1i64),
            _ => (args[0].as_f64() as i64, args[1].as_f64() as i64, {
                let s = args[2].as_f64() as i64;
                if s == 0 { 1 } else { s }
            }),
        };
        let mut result = Vec::new();
        if step > 0 {
            let mut i = start;
            while i < stop { result.push(Value::I32(i as i32)); i += step; }
        } else {
            let mut i = start;
            while i > stop { result.push(Value::I32(i as i32)); i += step; }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(result))))
    }));

    // enumerate(iterable) → array of [index, value] pairs
    vm.register_host_fn("vybe:array", "enumerate", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                let pairs: Vec<Value> = elems.iter().enumerate().map(|(i, v)| {
                    Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::I32(i as i32), v.clone()]))))
                }).collect();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // zip(a, b) → array of [a[i], b[i]] pairs
    vm.register_host_fn("vybe:array", "zip", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let get_elems = |v: &Value| -> Vec<Value> {
            if let Value::Object(obj) = v {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(e) = &o.kind { return e.clone(); }
            }
            vec![]
        };
        let a = args.first().map(get_elems).unwrap_or_default();
        let b = args.get(1).map(get_elems).unwrap_or_default();
        let len = a.len().min(b.len());
        let pairs: Vec<Value> = (0..len).map(|i| {
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![a[i].clone(), b[i].clone()]))))
        }).collect();
        Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))))
    }));

    // sorted(iterable) → new sorted array
    vm.register_host_fn("vybe:array", "sorted", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                let mut sorted = elems.clone();
                sorted.sort_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal));
                return Value::Object(Arc::new(Mutex::new(Object::new_array(sorted))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // reversed(iterable) → new reversed array
    vm.register_host_fn("vybe:array", "reversed", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                let mut rev = elems.clone();
                rev.reverse();
                return Value::Object(Arc::new(Mutex::new(Object::new_array(rev))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // sum(iterable)
    vm.register_host_fn("vybe:array", "sum", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                let total: f64 = elems.iter().map(|v| v.as_f64()).sum();
                return Value::F64(total);
            }
        }
        Value::F64(0.0)
    }));

    // min(iterable) or min(a, b, ...)
    vm.register_host_fn("vybe:array", "pymin", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if args.len() == 1 {
            if let Value::Object(obj) = &args[0] {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(elems) = &o.kind {
                    return elems.iter().min_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal)).cloned().unwrap_or(Value::Null);
                }
            }
        }
        args.iter().min_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal)).cloned().unwrap_or(Value::Null)
    }));

    // max(iterable) or max(a, b, ...)
    vm.register_host_fn("vybe:array", "pymax", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if args.len() == 1 {
            if let Value::Object(obj) = &args[0] {
                let o = obj.lock().unwrap();
                if let ObjectKind::Array(elems) = &o.kind {
                    return elems.iter().max_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal)).cloned().unwrap_or(Value::Null);
                }
            }
        }
        args.iter().max_by(|a, b| a.as_f64().partial_cmp(&b.as_f64()).unwrap_or(std::cmp::Ordering::Equal)).cloned().unwrap_or(Value::Null)
    }));

    // any(iterable) → bool
    vm.register_host_fn("vybe:array", "any", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                return Value::Bool(elems.iter().any(|v| is_truthy(v)));
            }
        }
        Value::Bool(false)
    }));

    // all(iterable) → bool
    vm.register_host_fn("vybe:array", "all", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                return Value::Bool(elems.iter().all(|v| is_truthy(v)));
            }
        }
        Value::Bool(true)
    }));

    // dict_items(dict) → array of [key, value] pairs
    vm.register_host_fn("vybe:array", "dictItems", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let pairs: Vec<Value> = o.properties.iter()
                .filter(|(k, _)| !k.starts_with("__"))
                .map(|(k, v)| {
                    Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                        Value::String(Arc::from(k.as_str())), v.clone()
                    ]))))
                }).collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // str_contains(haystack, needle) → bool (for "x in string")
    vm.register_host_fn("vybe:array", "strContains", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let haystack = args.first().map(|v| v.to_string()).unwrap_or_default();
        let needle = args.get(1).map(|v| v.to_string()).unwrap_or_default();
        Value::Bool(haystack.contains(&needle))
    }));

    // dict_contains(dict, key) → bool (for "key in dict")
    vm.register_host_fn("vybe:array", "dictContains", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| v.to_string()).unwrap_or_default();
            let o = obj.lock().unwrap();
            return Value::Bool(o.properties.contains_key(&key));
        }
        Value::Bool(false)
    }));

    // isinstance(obj, type_name_str) — simplified
    vm.register_host_fn("vybe:array", "isinstance", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let type_name = args.get(1).map(|v| v.to_string()).unwrap_or_default().to_lowercase();
        let val = args.first().unwrap_or(&Value::Null);
        let result = match (val, type_name.as_str()) {
            (Value::I32(_) | Value::I64(_), "int") => true,
            (Value::F64(_), "float") => true,
            (Value::String(_), "str") => true,
            (Value::Bool(_), "bool") => true,
            (Value::Null, "nonetype") => true,
            (Value::Object(o), "list") => matches!(o.lock().unwrap().kind, ObjectKind::Array(_)),
            (Value::Object(o), "dict") => !matches!(o.lock().unwrap().kind, ObjectKind::Array(_)),
            _ => false,
        };
        Value::Bool(result)
    }));

    // type(obj) → string name
    vm.register_host_fn("vybe:array", "pytype", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let val = args.first().unwrap_or(&Value::Null);
        let name = match val {
            Value::I32(_) | Value::I64(_) => "int",
            Value::F64(_) => "float",
            Value::String(_) => "str",
            Value::Bool(_) => "bool",
            Value::Null => "NoneType",
            Value::Object(o) => {
                let o = o.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(_) => "list",
                    ObjectKind::Function(_) => "function",
                    _ => "object",
                }
            }
            _ => "object",
        };
        Value::String(Arc::from(name))
    }));

    // list(iterable) — convert to list (copy)
    vm.register_host_fn("vybe:array", "list", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
            }
            // Dict → list of keys
            let keys: Vec<Value> = o.properties.keys().map(|k| Value::String(Arc::from(k.as_str()))).collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(keys))));
        }
        if let Some(Value::String(s)) = args.first() {
            let chars: Vec<Value> = s.chars().map(|c| Value::String(Arc::from(c.to_string().as_str()))).collect();
            return Value::Object(Arc::new(Mutex::new(Object::new_array(chars))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // dict() — create empty dict
    vm.register_host_fn("vybe:array", "dict", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Object(Arc::new(Mutex::new(Object::new())))
    }));

    // set() — create set as array (simplified)
    vm.register_host_fn("vybe:array", "pyset", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                let mut seen = std::collections::HashSet::new();
                let mut result = Vec::new();
                for e in elems {
                    let s = e.to_string();
                    if seen.insert(s) { result.push(e.clone()); }
                }
                return Value::Object(Arc::new(Mutex::new(Object::new_array(result))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // tuple() — same as list for our purposes
    vm.register_host_fn("vybe:array", "tuple", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &o.kind {
                return Value::Object(Arc::new(Mutex::new(Object::new_array(elems.clone()))));
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    // fill(arr, val, start?, end?) → arr (mutated in-place)
    // JS semantics: fill indices [start, end) with val. Returns arr.
    vm.register_host_fn("vybe:array", "fill", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let val = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                let len = elems.len() as i64;
                let start = args.get(2).map(|v| v.as_f64() as i64).unwrap_or(0);
                let end = args.get(3).map(|v| v.as_f64() as i64).unwrap_or(len);
                let s = if start < 0 { (len + start).max(0) as usize } else { start.min(len) as usize };
                let e = if end < 0 { (len + end).max(0) as usize } else { end.min(len) as usize };
                for i in s..e { elems[i] = val.clone(); }
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));
}

fn norm(args: &[Value], idx: usize, default: i64, len: i64) -> usize {
    let v = args.get(idx).map(|v| v.as_f64() as i64).unwrap_or(default);
    if v < 0 { (len + v).max(0) as usize } else { v.min(len) as usize }
}

fn is_truthy(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::F64(f) => *f != 0.0,
        Value::String(s) => !s.is_empty(),
        Value::Object(o) => {
            let o = o.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(a) => !a.is_empty(),
                _ => true,
            }
        }
        _ => true,
    }
}
