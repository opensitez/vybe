use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::ObjectKind;

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:json", "stringify", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::String(Rc::from(stringify(args.first().unwrap_or(&Value::Null)).as_str()))
    }));

    vm.register_host_fn("vybe:json", "parse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        parse_json(&s)
    }));
}

use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::value::Object;

fn parse_json(s: &str) -> Value {
    let s = s.trim();
    if s == "null" { return Value::Null; }
    if s == "true" { return Value::Bool(true); }
    if s == "false" { return Value::Bool(false); }
    if let Ok(n) = s.parse::<f64>() { return Value::F64(n); }
    // String: "..."
    if s.starts_with('"') && s.ends_with('"') && s.len() >= 2 {
        let inner = &s[1..s.len()-1];
        let unescaped = inner
            .replace("\\\"", "\"")
            .replace("\\\\", "\\")
            .replace("\\n", "\n")
            .replace("\\t", "\t")
            .replace("\\r", "\r");
        return Value::String(Rc::from(unescaped.as_str()));
    }
    // Array: [...]
    if s.starts_with('[') && s.ends_with(']') {
        let inner = &s[1..s.len()-1].trim();
        if inner.is_empty() {
            return Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))));
        }
        let items = split_json_values(inner);
        let elems: Vec<Value> = items.iter().map(|item| parse_json(item.trim())).collect();
        return Value::Object(Rc::new(RefCell::new(Object::new_array(elems))));
    }
    // Object: {...}
    if s.starts_with('{') && s.ends_with('}') {
        let inner = &s[1..s.len()-1].trim();
        if inner.is_empty() {
            return Value::Object(Rc::new(RefCell::new(Object::new())));
        }
        let mut obj = Object::new();
        let pairs = split_json_values(inner);
        for pair in &pairs {
            let pair = pair.trim();
            if let Some(colon_pos) = find_colon(pair) {
                let key = pair[..colon_pos].trim();
                let val = pair[colon_pos+1..].trim();
                // Remove quotes from key
                let key = if key.starts_with('"') && key.ends_with('"') {
                    &key[1..key.len()-1]
                } else { key };
                obj.properties.insert(key.to_string(), parse_json(val));
            }
        }
        return Value::Object(Rc::new(RefCell::new(obj)));
    }
    // Unknown — return as string
    Value::String(Rc::from(s))
}

/// Split comma-separated JSON values, respecting nesting.
fn split_json_values(s: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut current = String::new();
    let mut depth = 0i32;
    let mut in_string = false;
    let mut escape = false;

    for ch in s.chars() {
        if escape {
            current.push(ch);
            escape = false;
            continue;
        }
        if ch == '\\' && in_string {
            current.push(ch);
            escape = true;
            continue;
        }
        if ch == '"' {
            in_string = !in_string;
            current.push(ch);
            continue;
        }
        if in_string {
            current.push(ch);
            continue;
        }
        match ch {
            '[' | '{' => { depth += 1; current.push(ch); }
            ']' | '}' => { depth -= 1; current.push(ch); }
            ',' if depth == 0 => {
                result.push(current.clone());
                current.clear();
            }
            _ => current.push(ch),
        }
    }
    if !current.is_empty() {
        result.push(current);
    }
    result
}

/// Find the colon separating key:value in a JSON pair, respecting strings.
fn find_colon(s: &str) -> Option<usize> {
    let mut in_string = false;
    let mut escape = false;
    for (i, ch) in s.chars().enumerate() {
        if escape { escape = false; continue; }
        if ch == '\\' { escape = true; continue; }
        if ch == '"' { in_string = !in_string; continue; }
        if !in_string && ch == ':' { return Some(i); }
    }
    None
}

fn stringify(v: &Value) -> String {
    match v {
        Value::Null | Value::Undefined => "null".into(),
        Value::Bool(b) => b.to_string(),
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => {
            if n.is_nan() || n.is_infinite() { "null".into() }
            else if *n == (*n as i64) as f64 { format!("{}", *n as i64) }
            else { format!("{}", n) }
        }
        Value::String(s) => format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\"").replace('\n', "\\n")),
        Value::Object(obj) => {
            let o = obj.borrow();
            match &o.kind {
                ObjectKind::Array(elems) => {
                    let parts: Vec<String> = elems.iter().map(|e| stringify(e)).collect();
                    format!("[{}]", parts.join(","))
                }
                _ => {
                    let parts: Vec<String> = o.properties.iter()
                        .map(|(k, v)| format!("\"{}\":{}", k, stringify(v)))
                        .collect();
                    format!("{{{}}}", parts.join(","))
                }
            }
        }
        Value::V128(b) => {
            let parts: Vec<String> = b.iter().map(|x| x.to_string()).collect();
            format!("[{}]", parts.join(","))
        }
        Value::WeakRef(_) => "null".into(),
    }
}
