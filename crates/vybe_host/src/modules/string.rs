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

    // charCodeAt(str, index) → number
    vm.register_host_fn("vybe:string", "charCodeAt", Box::new(|a| {
        let st = s(a, 0);
        let idx = f(a, 1) as usize;
        match st.chars().nth(idx) {
            Some(c) => Value::F64(c as u32 as f64),
            None => Value::F64(f64::NAN),
        }
    }));

    // fromCharCode(code, code, ...) → string
    vm.register_host_fn("vybe:string", "fromCharCode", Box::new(|a| {
        let result: String = a.iter()
            .map(|v| char::from_u32(v.as_f64() as u32).unwrap_or('\0'))
            .collect();
        Value::String(Rc::from(result.as_str()))
    }));

    // repeat(str, count) → string
    vm.register_host_fn("vybe:string", "repeat", Box::new(|a| {
        let st = s(a, 0);
        let count = f(a, 1) as usize;
        Value::String(Rc::from(st.repeat(count).as_str()))
    }));

    // padStart(str, targetLength, padString?) → string
    vm.register_host_fn("vybe:string", "padStart", Box::new(|a| {
        let st = s(a, 0);
        let target = f(a, 1) as usize;
        let pad = if a.len() > 2 { s(a, 2) } else { " ".into() };
        if st.len() >= target { return Value::String(Rc::from(st.as_str())); }
        let needed = target - st.len();
        let padding: String = pad.chars().cycle().take(needed).collect();
        Value::String(Rc::from(format!("{}{}", padding, st).as_str()))
    }));

    // padEnd(str, targetLength, padString?) → string
    vm.register_host_fn("vybe:string", "padEnd", Box::new(|a| {
        let st = s(a, 0);
        let target = f(a, 1) as usize;
        let pad = if a.len() > 2 { s(a, 2) } else { " ".into() };
        if st.len() >= target { return Value::String(Rc::from(st.as_str())); }
        let needed = target - st.len();
        let padding: String = pad.chars().cycle().take(needed).collect();
        Value::String(Rc::from(format!("{}{}", st, padding).as_str()))
    }));

    // replaceAll(str, search, replace) → string
    vm.register_host_fn("vybe:string", "replaceAll", Box::new(|a| {
        Value::String(Rc::from(s(a, 0).replace(&s(a, 1), &s(a, 2)).as_str()))
    }));

    // trimStart / trimEnd
    vm.register_host_fn("vybe:string", "trimStart", Box::new(|a| Value::String(Rc::from(s(a, 0).trim_start()))));
    vm.register_host_fn("vybe:string", "trimEnd",   Box::new(|a| Value::String(Rc::from(s(a, 0).trim_end()))));

    // --- VB-compatible string functions (available to all languages) ---

    // left(str, n) → first n characters
    vm.register_host_fn("vybe:string", "left", Box::new(|a| {
        let st = s(a, 0);
        let n = f(a, 1) as usize;
        let end = n.min(st.len());
        Value::String(Rc::from(&st[..end]))
    }));

    // right(str, n) → last n characters
    vm.register_host_fn("vybe:string", "right", Box::new(|a| {
        let st = s(a, 0);
        let n = f(a, 1) as usize;
        let start = st.len().saturating_sub(n);
        Value::String(Rc::from(&st[start..]))
    }));

    // mid(str, start, length?) → substring (1-based start like VB)
    vm.register_host_fn("vybe:string", "mid", Box::new(|a| {
        let st = s(a, 0);
        let start = ((f(a, 1) as usize).saturating_sub(1)).min(st.len());
        let end = if a.len() > 2 {
            (start + f(a, 2) as usize).min(st.len())
        } else {
            st.len()
        };
        Value::String(Rc::from(&st[start..end]))
    }));

    // instr(str, search) → 1-based position, 0 if not found
    // instr(start, str, search) → 1-based position starting from start
    vm.register_host_fn("vybe:string", "instr", Box::new(|a| {
        if a.len() >= 3 {
            // instr(start, str, search)
            let start = (f(a, 0) as usize).saturating_sub(1);
            let st = s(a, 1);
            let search = s(a, 2);
            match st[start..].find(&search) {
                Some(idx) => Value::F64((start + idx + 1) as f64),
                None => Value::F64(0.0),
            }
        } else {
            // instr(str, search)
            let st = s(a, 0);
            let search = s(a, 1);
            match st.find(&search) {
                Some(idx) => Value::F64((idx + 1) as f64),
                None => Value::F64(0.0),
            }
        }
    }));

    // asc(str) → char code of first character
    vm.register_host_fn("vybe:string", "asc", Box::new(|a| {
        match s(a, 0).chars().next() {
            Some(c) => Value::F64(c as u32 as f64),
            None => Value::F64(0.0),
        }
    }));

    // chr(code) → single character string
    vm.register_host_fn("vybe:string", "chr", Box::new(|a| {
        match char::from_u32(f(a, 0) as u32) {
            Some(c) => Value::String(Rc::from(c.to_string().as_str())),
            None => Value::String(Rc::from("")),
        }
    }));

    // space(n) → string of n spaces
    vm.register_host_fn("vybe:string", "space", Box::new(|a| {
        Value::String(Rc::from(" ".repeat(f(a, 0) as usize).as_str()))
    }));

    // string(n, char) → string of n copies of char
    vm.register_host_fn("vybe:string", "stringRepeat", Box::new(|a| {
        let n = f(a, 0) as usize;
        let ch = s(a, 1);
        let c = ch.chars().next().unwrap_or(' ');
        Value::String(Rc::from(c.to_string().repeat(n).as_str()))
    }));

    // lcase/ucase aliases (same as toLowerCase/toUpperCase but available by VB name)
    vm.register_host_fn("vybe:string", "lcase", Box::new(|a| Value::String(Rc::from(s(a, 0).to_lowercase().as_str()))));
    vm.register_host_fn("vybe:string", "ucase", Box::new(|a| Value::String(Rc::from(s(a, 0).to_uppercase().as_str()))));
    vm.register_host_fn("vybe:string", "ltrim", Box::new(|a| Value::String(Rc::from(s(a, 0).trim_start()))));
    vm.register_host_fn("vybe:string", "rtrim", Box::new(|a| Value::String(Rc::from(s(a, 0).trim_end()))));
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
