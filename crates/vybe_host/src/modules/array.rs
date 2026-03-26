use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:array", "push", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                for arg in &args[1..] { elems.push(arg.clone()); }
                let len = elems.len() as f64;
                drop(o);
                obj.borrow_mut().properties.insert("length".into(), Value::F64(len));
                return Value::F64(len);
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:array", "pop", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                let val = elems.pop().unwrap_or(Value::Null);
                let len = elems.len() as f64;
                drop(o);
                obj.borrow_mut().properties.insert("length".into(), Value::F64(len));
                return val;
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:array", "shift", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind {
                if !elems.is_empty() {
                    let val = elems.remove(0);
                    let len = elems.len() as f64;
                    drop(o);
                    obj.borrow_mut().properties.insert("length".into(), Value::F64(len));
                    return val;
                }
            }
        }
        Value::Null
    }));
    vm.register_host_fn("vybe:array", "length", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                return Value::F64(elems.len() as f64);
            }
        }
        Value::F64(0.0)
    }));
    vm.register_host_fn("vybe:array", "join", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                let sep = if args.len() > 1 { format!("{}", args[1]) } else { ",".into() };
                let parts: Vec<String> = elems.iter().map(|v| format!("{}", v)).collect();
                return Value::String(Rc::from(parts.join(&sep).as_str()));
            }
        }
        Value::String(Rc::from(""))
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
    vm.register_host_fn("vybe:array", "reverse", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            if let ObjectKind::Array(ref mut elems) = o.kind { elems.reverse(); }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));
    vm.register_host_fn("vybe:array", "concat", Box::new(|args: &[Value]| {
        let mut result = Vec::new();
        for arg in args {
            if let Value::Object(obj) = arg {
                let o = obj.borrow();
                if let ObjectKind::Array(ref elems) = o.kind {
                    result.extend(elems.clone());
                    continue;
                }
            }
            result.push(arg.clone());
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(result))))
    }));
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
}

fn norm(args: &[Value], idx: usize, default: i64, len: i64) -> usize {
    let v = args.get(idx).map(|v| v.as_f64() as i64).unwrap_or(default);
    if v < 0 { (len + v).max(0) as usize } else { v.min(len) as usize }
}
