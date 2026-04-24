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

    vm.register_host_fn("vybe:array", "lastIndexOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                let search = args.get(1).cloned().unwrap_or(Value::Null);
                for (i, elem) in elems.iter().enumerate().rev() {
                    if elem.eq(&search) { return Value::F64(i as f64); }
                }
                return Value::F64(-1.0);
            }
        }
        // Fall back to string last index of via format
        if let (Some(Value::String(s)), Some(search)) = (args.first(), args.get(1)) {
            let needle = format!("{}", search);
            if let Some(pos) = s.rfind(needle.as_str()) {
                return Value::F64(pos as f64);
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
    vm.register_host_fn("vybe:js-math", "dynMul", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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

    // clear(collection) → null
    // Compatibility shim for profiles that still emit host:vybe:array:clear
    // for list/dict/array-style receivers.
    vm.register_host_fn("vybe:array", "clear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            match &mut o.kind {
                ObjectKind::Array(elems) => {
                    elems.clear();
                    o.properties.insert("length".into(), Value::F64(0.0));
                    o.properties.insert("count".into(), Value::F64(0.0));
                }
                _ => {
                    let data_obj = match o.properties.get("__data") {
                        Some(Value::Object(data)) => Some(data.clone()),
                        _ => None,
                    };
                    if let Some(data) = data_obj {
                        data.lock().unwrap().properties.clear();
                        o.properties.insert("count".into(), Value::F64(0.0));
                    }
                    let items_obj = match o.properties.get("__items") {
                        Some(Value::Object(items)) => Some(items.clone()),
                        _ => None,
                    };
                    if let Some(items) = items_obj {
                        let mut items_obj = items.lock().unwrap();
                        if let ObjectKind::Array(elems) = &mut items_obj.kind {
                            elems.clear();
                        }
                        drop(items_obj);
                        o.properties.insert("count".into(), Value::F64(0.0));
                    }
                }
            }
        }
        Value::Null
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
    // Array.of(...args) — creates array from variable arguments
    vm.register_host_fn("vybe:array", "of", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let mut arr = Object::new();
        arr.kind = ObjectKind::Array(args.to_vec());
        Value::Object(Arc::new(Mutex::new(arr)))
    }));

    // Array constructor: new Array() or new Array(5) or new Array("a", "b")
    // vybex `new X(args)` passes user args only (no this prefix).
    vm.register_host_fn("vybe:array", "arrayNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let elems: Vec<Value> = if args.len() == 1 {
            let v = &args[0];
            if matches!(v, Value::F64(_) | Value::I32(_) | Value::I64(_)) {
                let n = v.as_f64() as usize;
                vec![Value::Null; n]
            } else {
                vec![v.clone()]
            }
        } else {
            args.to_vec()
        };
        let mut obj = Object::new();
        obj.kind = ObjectKind::Array(elems);
        obj.properties.insert("__type".into(), Value::String(Arc::from("Array")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // arr.keys() → iterator of indices [0, 1, 2, ...] (returns array for spread support)
    vm.register_host_fn("vybe:array", "arrKeys", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let a = arr.lock().unwrap();
            if let ObjectKind::Array(elements) = &a.kind {
                let keys: Vec<Value> = (0..elements.len()).map(|i| Value::F64(i as f64)).collect();
                let mut out = Object::new();
                out.kind = ObjectKind::Array(keys);
                return Value::Object(Arc::new(Mutex::new(out)));
            }
        }
        let mut out = Object::new();
        out.kind = ObjectKind::Array(vec![]);
        Value::Object(Arc::new(Mutex::new(out)))
    }));

    // arr.values() → iterator of values (returns shallow copy for spread support)
    vm.register_host_fn("vybe:array", "arrValues", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let a = arr.lock().unwrap();
            if let ObjectKind::Array(elements) = &a.kind {
                let mut out = Object::new();
                out.kind = ObjectKind::Array(elements.clone());
                return Value::Object(Arc::new(Mutex::new(out)));
            }
        }
        let mut out = Object::new();
        out.kind = ObjectKind::Array(vec![]);
        Value::Object(Arc::new(Mutex::new(out)))
    }));

    // Array.copyWithin(arr, target, start, [end])
    vm.register_host_fn("vybe:array", "copyWithin", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(arr)) = args.first() {
            let mut a = arr.lock().unwrap();
            if let ObjectKind::Array(ref mut elements) = a.kind {
                let len = elements.len() as i64;
                let parse_i = |v: &Value, default: i64| -> i64 {
                    format!("{}", v).parse::<f64>().unwrap_or(default as f64) as i64
                };
                let target = args.get(1).map(|v| parse_i(v, 0)).unwrap_or(0);
                let start = args.get(2).map(|v| parse_i(v, 0)).unwrap_or(0);
                let end = args.get(3).map(|v| parse_i(v, len)).unwrap_or(len);
                let target = if target < 0 { (len + target).max(0) } else { target.min(len) } as usize;
                let start = if start < 0 { (len + start).max(0) } else { start.min(len) } as usize;
                let end = if end < 0 { (len + end).max(0) } else { end.min(len) } as usize;
                if start < end && target < elements.len() {
                    let to_copy: Vec<Value> = elements[start..end].to_vec();
                    let copy_len = to_copy.len().min(elements.len() - target);
                    for (i, v) in to_copy.into_iter().take(copy_len).enumerate() {
                        elements[target + i] = v;
                    }
                }
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // Array.at(arr, index) — supports negative indices for arrays AND strings
    vm.register_host_fn("vybe:array", "at", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.get(1).map(|v| format!("{}", v).parse::<f64>().unwrap_or(0.0) as i64).unwrap_or(0);
        match args.first() {
            Some(Value::Object(arr)) => {
                let a = arr.lock().unwrap();
                if let ObjectKind::Array(elements) = &a.kind {
                    let len = elements.len() as i64;
                    let real_idx = if idx < 0 { len + idx } else { idx };
                    if real_idx >= 0 && (real_idx as usize) < elements.len() {
                        return elements[real_idx as usize].clone();
                    }
                }
                Value::Undefined
            }
            Some(Value::String(s)) => {
                let chars: Vec<char> = s.chars().collect();
                let len = chars.len() as i64;
                let real_idx = if idx < 0 { len + idx } else { idx };
                if real_idx >= 0 && (real_idx as usize) < chars.len() {
                    Value::String(Arc::from(chars[real_idx as usize].to_string().as_str()))
                } else {
                    Value::Undefined
                }
            }
            _ => Value::Undefined,
        }
    }));

    vm.register_host_fn("vybe:array", "from", Box::new(|ctx: &mut HostContext, args: &[Value]| {
        let mut elements: Vec<Value> = Vec::new();
        match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                // Native array: copy elements
                if let ObjectKind::Array(ref elems) = o.kind {
                    elements = elems.clone();
                } else {
                    let type_name = o.properties.get("__type")
                        .map(|v| format!("{}", v))
                        .unwrap_or_default();
                    if type_name == "Set" || type_name == "WeakSet" {
                        if let Some(Value::Object(items)) = o.properties.get("__items") {
                            let it = items.lock().unwrap();
                            if let ObjectKind::Array(ref elems) = it.kind {
                                elements = elems.clone();
                            }
                        }
                    } else if type_name == "Map" || type_name == "WeakMap" {
                        // Map → array of [key, value] pairs (use __keys order)
                        if let Some(Value::Object(data)) = o.properties.get("__data") {
                            let d = data.lock().unwrap();
                            if let Some(Value::Object(keys_arr)) = o.properties.get("__keys") {
                                let k = keys_arr.lock().unwrap();
                                if let ObjectKind::Array(ref keys) = k.kind {
                                    for key_v in keys {
                                        let ks = format!("{}", key_v);
                                        let v = d.properties.get(&ks).cloned().unwrap_or(Value::Undefined);
                                        let pair = vec![key_v.clone(), v];
                                        elements.push(Value::Object(Arc::new(Mutex::new(Object::new_array(pair)))));
                                    }
                                }
                            } else {
                                for (k, v) in &d.properties {
                                    let pair = vec![
                                        Value::String(Arc::from(k.as_str())),
                                        v.clone(),
                                    ];
                                    elements.push(Value::Object(Arc::new(Mutex::new(Object::new_array(pair)))));
                                }
                            }
                        }
                    } else {
                        // Array-like: object with `length` property and integer keys
                        if let Some(len_val) = o.properties.get("length") {
                            let len = len_val.as_f64() as usize;
                            for i in 0..len {
                                let key = i.to_string();
                                let v = o.properties.get(&key).cloned().unwrap_or(Value::Undefined);
                                elements.push(v);
                            }
                        }
                    }
                }
            }
            Some(Value::String(s)) => {
                // Array.from(string) → array of single-char strings
                for c in s.chars() {
                    elements.push(Value::String(Arc::from(c.to_string().as_str())));
                }
            }
            _ => {}
        }
        // Optional mapper: Array.from(src, (v, i) => ...)
        if let Some(mapper) = args.get(1) {
            if !matches!(mapper, Value::Null | Value::Undefined) {
                let mut mapped = Vec::with_capacity(elements.len());
                for (i, v) in elements.iter().enumerate() {
                    mapped.push(ctx.invoke(mapper, &[v.clone(), Value::F64(i as f64)]));
                }
                elements = mapped;
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
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
