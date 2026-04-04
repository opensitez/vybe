use std::rc::Rc;
use vybe_bytecode::{VM, Value, HostContext};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:string", "slice", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let len = st.len() as i64;
        let start = norm(f(a, 1) as i64, len);
        let end = if a.len() > 2 { norm(f(a, 2) as i64, len) } else { len as usize };
        if start < end { Value::String(Rc::from(&st[start..end])) }
        else { Value::String(Rc::from("")) }
    }));
    vm.register_host_fn("vybe:string", "indexOf", Box::new(|_ctx, a| {
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
    vm.register_host_fn("vybe:string", "includes",    Box::new(|_ctx, a| Value::Bool(s(a, 0).contains(&s(a, 1)))));
    vm.register_host_fn("vybe:string", "split", Box::new(|_ctx, a| {
        let parts: Vec<Value> = s(a, 0).split(&s(a, 1)).map(|p| Value::String(Rc::from(p))).collect();
        Value::Object(Rc::new(std::cell::RefCell::new(vybe_bytecode::value::Object::new_array(parts))))
    }));
    vm.register_host_fn("vybe:string", "replace",    Box::new(|_ctx, a| Value::String(Rc::from(s(a, 0).replacen(&s(a, 1), &s(a, 2), 1).as_str()))));
    vm.register_host_fn("vybe:string", "startsWith", Box::new(|_ctx, a| Value::Bool(s(a, 0).starts_with(&s(a, 1)))));
    vm.register_host_fn("vybe:string", "endsWith",   Box::new(|_ctx, a| Value::Bool(s(a, 0).ends_with(&s(a, 1)))));
    vm.register_host_fn("vybe:string", "charAt", Box::new(|_ctx, a| {
        match s(a, 0).chars().nth(f(a, 1) as usize) {
            Some(c) => Value::String(Rc::from(c.to_string().as_str())),
            None => Value::String(Rc::from("")),
        }
    }));
    vm.register_host_fn("vybe:string", "substring", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let start = (f(a, 1) as usize).min(st.len());
        let end = if a.len() > 2 { (f(a, 2) as usize).min(st.len()) } else { st.len() };
        let (start, end) = if start > end { (end, start) } else { (start, end) };
        Value::String(Rc::from(&st[start..end]))
    }));

    // charCodeAt(str, index) → number
    vm.register_host_fn("vybe:string", "charCodeAt", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let idx = f(a, 1) as usize;
        match st.chars().nth(idx) {
            Some(c) => Value::F64(c as u32 as f64),
            None => Value::F64(f64::NAN),
        }
    }));

    // fromCharCode(code, code, ...) → string
    vm.register_host_fn("vybe:string", "fromCharCode", Box::new(|_ctx, a| {
        let result: String = a.iter()
            .map(|v| char::from_u32(v.as_f64() as u32).unwrap_or('\0'))
            .collect();
        Value::String(Rc::from(result.as_str()))
    }));

    // repeat(str, count) → string
    vm.register_host_fn("vybe:string", "repeat", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let count = f(a, 1) as usize;
        Value::String(Rc::from(st.repeat(count).as_str()))
    }));

    // replaceAll(str, search, replace) → string
    vm.register_host_fn("vybe:string", "replaceAll", Box::new(|_ctx, a| {
        Value::String(Rc::from(s(a, 0).replace(&s(a, 1), &s(a, 2)).as_str()))
    }));

    // --- VB-compatible string functions (available to all languages) ---

    // left(str, n) → first n characters
    vm.register_host_fn("vybe:string", "left", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let n = f(a, 1) as usize;
        let end = n.min(st.len());
        Value::String(Rc::from(&st[..end]))
    }));

    // right(str, n) → last n characters
    vm.register_host_fn("vybe:string", "right", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let n = f(a, 1) as usize;
        let start = st.len().saturating_sub(n);
        Value::String(Rc::from(&st[start..]))
    }));

    // mid(str, start, length?) → substring (1-based start like VB)
    vm.register_host_fn("vybe:string", "mid", Box::new(|_ctx, a| {
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
    vm.register_host_fn("vybe:string", "instr", Box::new(|_ctx, a| {
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

    // string(n, char) → string of n copies of char
    vm.register_host_fn("vybe:string", "stringRepeat", Box::new(|_ctx, a| {
        let n = f(a, 0) as usize;
        let ch = s(a, 1);
        let c = ch.chars().next().unwrap_or(' ');
        Value::String(Rc::from(c.to_string().repeat(n).as_str()))
    }));

    // instrrev(str, search) → 1-based position of LAST occurrence, 0 if not found
    vm.register_host_fn("vybe:string", "instrrev", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let search = s(a, 1);
        match st.rfind(&search) {
            Some(idx) => Value::F64((idx + 1) as f64),
            None => Value::F64(0.0),
        }
    }));

    // format(value, formatStr) → formatted string
    // Supports both VB6 Format(value, spec) and .NET String.Format("{0}...", args...)
    vm.register_host_fn("vybe:string", "format", Box::new(|_ctx, a| {
        let first = s(a, 0);
        // Detect .NET composite format: first arg is a string containing {0}, {1}, etc.
        if first.contains("{0}") || first.contains("{1}") || first.contains("{2}") {
            let mut result = first.clone();
            for (i, arg) in a[1..].iter().enumerate() {
                let placeholder = format!("{{{}}}", i);
                result = result.replace(&placeholder, &format!("{}", arg));
            }
            Value::String(Rc::from(result.as_str()))
        } else {
            // VB6 Format(value, formatSpec)
            let val = f(a, 0);
            let fmt = s(a, 1).to_lowercase();
            let result = match fmt.as_str() {
                "0" | "0.0" | "#.#" | "fixed" => format!("{:.1}", val),
                "0.00" | "#.##" | "standard" => format!("{:.2}", val),
                "percent" => format!("{:.2}%", val * 100.0),
                "currency" => format!("${:.2}", val),
                "scientific" => format!("{:e}", val),
                "yes/no" => if val != 0.0 { "Yes".into() } else { "No".into() },
                "true/false" => if val != 0.0 { "True".into() } else { "False".into() },
                "on/off" => if val != 0.0 { "On".into() } else { "Off".into() },
                _ => format!("{}", val),
            };
            Value::String(Rc::from(result.as_str()))
        }
    }));

    // lset(str, length) → left-align in field
    vm.register_host_fn("vybe:string", "lset", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let len = f(a, 1) as usize;
        Value::String(Rc::from(format!("{:<width$}", st, width = len).as_str()))
    }));

    // rset(str, length) → right-align in field
    vm.register_host_fn("vybe:string", "rset", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let len = f(a, 1) as usize;
        Value::String(Rc::from(format!("{:>width$}", st, width = len).as_str()))
    }));

    // filter(arr, match, include?) → filtered array
    vm.register_host_fn("vybe:string", "filter", Box::new(|_ctx, a| {
        use std::cell::RefCell;
        use vybe_bytecode::value::{Object, ObjectKind};
        let match_str = s(a, 1);
        let include = if a.len() > 2 { f(a, 2) != 0.0 } else { true };
        let mut results = Vec::new();
        if let Some(Value::Object(obj)) = a.first() {
            let o = obj.borrow();
            if let ObjectKind::Array(ref elems) = o.kind {
                for elem in elems {
                    let es = format!("{}", elem);
                    let contains = es.contains(&match_str);
                    if (include && contains) || (!include && !contains) {
                        results.push(elem.clone());
                    }
                }
            }
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(results))))
    }));

    // count(str, sub) → number of non-overlapping occurrences
    vm.register_host_fn("vybe:string", "count", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let haystack = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let needle = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if needle.is_empty() {
            return Value::I32(0);
        }
        Value::I32(haystack.matches(&needle).count() as i32)
    }));

    // padStart(str, width) — zero-fill / right-justify
    vm.register_host_fn("vybe:string", "padStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let width = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let fill = args.get(2).map(|v| format!("{}", v)).unwrap_or_else(|| " ".to_string());
        let fill_char = fill.chars().next().unwrap_or(' ');
        if s.len() >= width {
            Value::String(Rc::from(s))
        } else {
            let padding: String = std::iter::repeat(fill_char).take(width - s.len()).collect();
            Value::String(Rc::from(format!("{}{}", padding, s)))
        }
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
