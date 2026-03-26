use std::rc::Rc;
use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:string", "length", Box::new(|a| {
        Value::F64(s(a, 0).len() as f64)
    }));
    vm.register_host_fn("vybe:string", "slice", Box::new(|a| {
        let st = s(a, 0);
        let len = st.len() as i64;
        let start = norm(f(a, 1) as i64, len);
        let end = if a.len() > 2 { norm(f(a, 2) as i64, len) } else { len as usize };
        if start < end { Value::String(Rc::from(&st[start..end])) }
        else { Value::String(Rc::from("")) }
    }));
    vm.register_host_fn("vybe:string", "indexOf", Box::new(|a| {
        match a.first() {
            Some(Value::String(st)) => {
                match st.find(&s(a, 1)) {
                    Some(idx) => Value::F64(idx as f64),
                    None => Value::F64(-1.0),
                }
            }
            Some(Value::Object(obj)) => {
                let o = obj.borrow();
                if let vybe_bytecode::value::ObjectKind::Array(ref elems) = o.kind {
                    let search = a.get(1).cloned().unwrap_or(Value::Null);
                    for (i, elem) in elems.iter().enumerate() {
                        if elem.eq(&search) { return Value::F64(i as f64); }
                    }
                }
                Value::F64(-1.0)
            }
            _ => Value::F64(-1.0),
        }
    }));
    vm.register_host_fn("vybe:string", "includes",    Box::new(|a| Value::Bool(s(a, 0).contains(&s(a, 1)))));
    vm.register_host_fn("vybe:string", "toUpperCase", Box::new(|a| Value::String(Rc::from(s(a, 0).to_uppercase().as_str()))));
    vm.register_host_fn("vybe:string", "toLowerCase", Box::new(|a| Value::String(Rc::from(s(a, 0).to_lowercase().as_str()))));
    vm.register_host_fn("vybe:string", "trim",        Box::new(|a| Value::String(Rc::from(s(a, 0).trim()))));
    vm.register_host_fn("vybe:string", "split", Box::new(|a| {
        let parts: Vec<Value> = s(a, 0).split(&s(a, 1)).map(|p| Value::String(Rc::from(p))).collect();
        Value::Object(Rc::new(std::cell::RefCell::new(vybe_bytecode::value::Object::new_array(parts))))
    }));
    vm.register_host_fn("vybe:string", "replace",    Box::new(|a| Value::String(Rc::from(s(a, 0).replacen(&s(a, 1), &s(a, 2), 1).as_str()))));
    vm.register_host_fn("vybe:string", "startsWith", Box::new(|a| Value::Bool(s(a, 0).starts_with(&s(a, 1)))));
    vm.register_host_fn("vybe:string", "endsWith",   Box::new(|a| Value::Bool(s(a, 0).ends_with(&s(a, 1)))));
    vm.register_host_fn("vybe:string", "charAt", Box::new(|a| {
        match s(a, 0).chars().nth(f(a, 1) as usize) {
            Some(c) => Value::String(Rc::from(c.to_string().as_str())),
            None => Value::String(Rc::from("")),
        }
    }));
    vm.register_host_fn("vybe:string", "substring", Box::new(|a| {
        let st = s(a, 0);
        let start = (f(a, 1) as usize).min(st.len());
        let end = if a.len() > 2 { (f(a, 2) as usize).min(st.len()) } else { st.len() };
        let (start, end) = if start > end { (end, start) } else { (start, end) };
        Value::String(Rc::from(&st[start..end]))
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}
fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}
fn norm(idx: i64, len: i64) -> usize {
    if idx < 0 { (len + idx).max(0) as usize } else { idx.min(len) as usize }
}
