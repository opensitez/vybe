//! `node:querystring` — Node.js URL query-string parser/serializer.
//!
//! Reference: <https://nodejs.org/api/querystring.html>.

use std::collections::HashMap;
use std::sync::Arc;
use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn qs_escape(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    for b in input.bytes() {
        match b {
            b'A'..=b'Z'
            | b'a'..=b'z'
            | b'0'..=b'9'
            | b'-'
            | b'_'
            | b'.'
            | b'~'
            | b'!'
            | b'\''
            | b'('
            | b')'
            | b'*' => {
                out.push(b as char);
            }
            b' ' => out.push_str("%20"),
            _ => out.push_str(&format!("%{b:02X}")),
        }
    }
    out
}

fn qs_unescape(input: &str) -> String {
    // Collect all bytes first (handles multi-byte UTF-8 percent-encoded sequences)
    let mut bytes: Vec<u8> = Vec::with_capacity(input.len());
    let src = input.as_bytes();
    let mut i = 0;
    while i < src.len() {
        if src[i] == b'+' {
            bytes.push(b' ');
            i += 1;
        } else if src[i] == b'%' && i + 2 < src.len() {
            if let Ok(hex) = std::str::from_utf8(&src[i + 1..i + 3]) {
                if let Ok(b) = u8::from_str_radix(hex, 16) {
                    bytes.push(b);
                    i += 3;
                    continue;
                }
            }
            bytes.push(src[i]);
            i += 1;
        } else {
            bytes.push(src[i]);
            i += 1;
        }
    }
    String::from_utf8_lossy(&bytes).into_owned()
}

fn qs_parse(input: &str, sep: char, eq: char) -> Value {
    let mut map: HashMap<String, Vec<String>> = HashMap::new();
    for pair in input.split(sep) {
        if pair.is_empty() {
            continue;
        }
        let (k, v) = if let Some(pos) = pair.find(eq) {
            (&pair[..pos], &pair[pos + 1..])
        } else {
            (pair, "")
        };
        let key = qs_unescape(k);
        let val = qs_unescape(v);
        map.entry(key).or_default().push(val);
    }

    let mut obj = Object::new();
    for (k, mut vals) in map {
        if vals.len() == 1 {
            obj.properties.insert(k, s(&vals.remove(0)));
        } else {
            let arr: Vec<Value> = vals.into_iter().map(|v| s(&v)).collect();
            obj.properties.insert(
                k,
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(arr))),
            );
        }
    }
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn qs_stringify(obj: &Value, sep: char, eq: char) -> String {
    let Value::Object(o) = obj else {
        return String::new();
    };
    let o = o.lock().unwrap();
    let mut parts = Vec::new();
    for (k, v) in &o.properties {
        if k.starts_with("__") {
            continue;
        }
        let encoded_k = qs_escape(k);
        match v {
            Value::String(s) => parts.push(format!("{}{}{}", encoded_k, eq, qs_escape(s))),
            Value::I32(n) => parts.push(format!("{}{}{}", encoded_k, eq, n)),
            Value::F64(f) => parts.push(format!("{}{}{}", encoded_k, eq, f)),
            Value::Bool(b) => parts.push(format!("{}{}{}", encoded_k, eq, b)),
            Value::Object(arr) => {
                let arr = arr.lock().unwrap();
                if let ObjectKind::Array(elems) = &arr.kind {
                    for elem in elems {
                        let val_str = match elem {
                            Value::String(s) => qs_escape(s),
                            Value::I32(n) => n.to_string(),
                            Value::F64(f) => f.to_string(),
                            _ => String::new(),
                        };
                        parts.push(format!("{}{}{}", encoded_k, eq, val_str));
                    }
                }
            }
            _ => {}
        }
    }
    parts.join(&sep.to_string())
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:querystring",
        "parse",
        Box::new(|_ctx, args| {
            let input = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            let sep = match args.get(1) {
                Some(Value::String(s)) if !s.is_empty() => s.chars().next().unwrap_or('&'),
                _ => '&',
            };
            let eq = match args.get(2) {
                Some(Value::String(s)) if !s.is_empty() => s.chars().next().unwrap_or('='),
                _ => '=',
            };
            qs_parse(&input, sep, eq)
        }),
    );

    vm.register_host_fn(
        "node:querystring",
        "stringify",
        Box::new(|_ctx, args| {
            let obj = args.first().cloned().unwrap_or(Value::Undefined);
            let sep = match args.get(1) {
                Some(Value::String(s)) if !s.is_empty() => s.chars().next().unwrap_or('&'),
                _ => '&',
            };
            let eq = match args.get(2) {
                Some(Value::String(s)) if !s.is_empty() => s.chars().next().unwrap_or('='),
                _ => '=',
            };
            s(&qs_stringify(&obj, sep, eq))
        }),
    );

    vm.register_host_fn(
        "node:querystring",
        "escape",
        Box::new(|_ctx, args| {
            let input = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            s(&qs_escape(&input))
        }),
    );

    vm.register_host_fn(
        "node:querystring",
        "unescape",
        Box::new(|_ctx, args| {
            let input = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            s(&qs_unescape(&input))
        }),
    );
}

#[allow(dead_code)]
fn _force_use(_: ObjectKind) {}
