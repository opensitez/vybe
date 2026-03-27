use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // push, pop, shift, length, join, reverse, concat — removed (now direct VM opcodes)

    // redim(array, newSize, preserve) → resized array
    vm.register_host_fn("vybe:array", "redim", Box::new(|args: &[Value]| {
        let new_size = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let preserve = matches!(args.get(2), Some(Value::Bool(true)));
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
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
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![Value::Null; new_size]))))
    }));

    vm.register_host_fn("vybe:array", "indexOf", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
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

    vm.register_host_fn("vybe:array", "slice", Box::new(|args: &[Value]| {
        match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.borrow();
                if let ObjectKind::Array(ref elems) = o.kind {
                    let len = elems.len() as i64;
                    let start = norm(args, 1, 0, len);
                    let end = norm(args, 2, len, len);
                    let sliced = if start < end { elems[start..end].to_vec() } else { vec![] };
                    return Value::Object(Rc::new(RefCell::new(Object::new_array(sliced))));
                }
                Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
            }
            Some(Value::String(s)) => {
                let len = s.len() as i64;
                let start = norm(args, 1, 0, len);
                let end = norm(args, 2, len, len);
                if start < end { Value::String(Rc::from(&s[start..end])) }
                else { Value::String(Rc::from("")) }
            }
            _ => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
        }
    }));
    // setAt(arr, index, value) — set element at index
    vm.register_host_fn("vybe:array", "setAt", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let value = args.get(2).cloned().unwrap_or(Value::Null);
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                if idx < elems.len() {
                    elems[idx] = value;
                }
            }
        }
        Value::Null
    }));

    // fill(arr, value, start?, end?) → arr
    vm.register_host_fn("vybe:array", "fill", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let value = args.get(1).cloned().unwrap_or(Value::Null);
            let mut o = obj.borrow_mut();
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
    vm.register_host_fn("vybe:array", "flat", Box::new(|args: &[Value]| {
        let mut result = Vec::new();
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                for elem in elems {
                    if let Value::Object(inner) = elem {
                        let inner_obj = inner.borrow();
                        if let ObjectKind::Array(ref inner_elems) = inner_obj.kind {
                            result.extend(inner_elems.clone());
                            continue;
                        }
                    }
                    result.push(elem.clone());
                }
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(result))))
    }));

    // includes(arr, value) → bool
    vm.register_host_fn("vybe:array", "includes", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let search = args.get(1).cloned().unwrap_or(Value::Null);
            let search_str = format!("{}", search);
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Bool(elems.iter().any(|e| format!("{}", e) == search_str));
            }
        }
        Value::Bool(false)
    }));

    // Array.isArray(value)
    vm.register_host_fn("vybe:array", "isArray", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            return Value::Bool(matches!(o.kind, ObjectKind::Array(_)));
        }
        Value::Bool(false)
    }));

    // Array.from(arrayLike) — creates a new array from an iterable/array-like
    vm.register_host_fn("vybe:array", "from", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::Object(Rc::new(RefCell::new(Object::new_array(elems.clone()))));
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // --- VB-compatible array functions ---

    // ubound(arr) → last valid index (length - 1)
    vm.register_host_fn("vybe:array", "ubound", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::F64(elems.len() as f64 - 1.0);
            }
        }
        Value::F64(-1.0)
    }));

    // lbound(arr) → always 0
    vm.register_host_fn("vybe:array", "lbound", Box::new(|_| Value::F64(0.0)));

    // unshift(arr, value) → prepend element, return new length
    vm.register_host_fn("vybe:array", "unshift", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                for i in (1..args.len()).rev() {
                    elems.insert(0, args[i].clone());
                }
                let len = elems.len() as f64;
                drop(o);
                obj.borrow_mut().properties.insert("length".into(), Value::F64(len));
                return Value::F64(len);
            }
        }
        Value::Null
    }));

    // splice(arr, start, deleteCount, ...items) → removed elements
    vm.register_host_fn("vybe:array", "splice", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let start = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let delete_count = args.get(2).map(|v| v.as_f64() as usize).unwrap_or(0);
            let insert_items: Vec<Value> = args.iter().skip(3).cloned().collect();
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                let start = start.min(elems.len());
                let end = (start + delete_count).min(elems.len());
                let removed: Vec<Value> = elems.drain(start..end).collect();
                for (i, item) in insert_items.into_iter().enumerate() {
                    elems.insert(start + i, item);
                }
                let len = elems.len() as f64;
                drop(o);
                obj.borrow_mut().properties.insert("length".into(), Value::F64(len));
                return Value::Object(Rc::new(RefCell::new(Object::new_array(removed))));
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));
}

fn norm(args: &[Value], idx: usize, default: i64, len: i64) -> usize {
    let v = args.get(idx).map(|v| v.as_f64() as i64).unwrap_or(default);
    if v < 0 { (len + v).max(0) as usize } else { v.min(len) as usize }
}
