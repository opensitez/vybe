//! `node:url` — Node.js URL module (WHATWG + legacy API).
//!
//! Reference: <https://nodejs.org/api/url.html>.

use std::sync::{Arc, Mutex};
use url::Url;
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::VM;

fn s(text: impl AsRef<str>) -> Value {
    Value::String(Arc::from(text.as_ref()))
}

fn empty_array() -> Value {
    arr_val(vec![])
}

fn arr_val(elems: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object {
        kind: ObjectKind::Array(elems),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })))
}

fn parse_url(input: &str, base: Option<&str>) -> Option<Url> {
    if let Some(base) = base {
        if let Ok(base_url) = Url::parse(base) {
            return base_url.join(input).ok();
        }
    }
    Url::parse(input).ok()
}

fn port_str(url: &Url) -> String {
    match url.port() {
        Some(p) => p.to_string(),
        None => String::new(),
    }
}

fn host_with_port(url: &Url) -> String {
    match url.port() {
        Some(p) => format!("{}:{}", url.host_str().unwrap_or(""), p),
        None => url.host_str().unwrap_or("").to_string(),
    }
}

fn origin_str(url: &Url) -> String {
    let scheme = url.scheme();
    let host = url.host_str().unwrap_or("");
    let default_port = match scheme {
        "https" => Some(443u16),
        "http" => Some(80),
        _ => None,
    };
    match (url.port(), default_port) {
        (Some(p), Some(dp)) if p != dp => format!("{scheme}://{host}:{p}"),
        _ => format!("{scheme}://{host}"),
    }
}

fn build_url_obj(url: &Url) -> Value {
    let mut o = Object::new();
    let href = url.as_str().to_string();
    o.properties.insert("href".into(), s(&href));
    o.properties.insert("protocol".into(), s(format!("{}:", url.scheme())));
    o.properties.insert("host".into(), s(host_with_port(url)));
    o.properties.insert("hostname".into(), s(url.host_str().unwrap_or("")));
    o.properties.insert("port".into(), s(port_str(url)));
    o.properties.insert("pathname".into(), s(url.path()));
    let search = url.query().map(|q| format!("?{q}")).unwrap_or_default();
    o.properties.insert("search".into(), s(&search));
    let hash = url.fragment().map(|f| format!("#{f}")).unwrap_or_default();
    o.properties.insert("hash".into(), s(&hash));
    o.properties.insert("origin".into(), s(origin_str(url)));
    o.properties.insert("username".into(), s(url.username()));
    o.properties.insert("password".into(), s(url.password().unwrap_or("")));
    // searchParams placeholder
    let sp = build_search_params(url.query().unwrap_or(""));
    o.properties.insert("searchParams".into(), sp);
    Value::Object(Arc::new(Mutex::new(o)))
}

/// Parse a query string into parallel __keys/__vals arrays stored on an Object.
fn build_search_params(query: &str) -> Value {
    let mut keys: Vec<Value> = vec![];
    let mut vals: Vec<Value> = vec![];
    for pair in query.split('&') {
        if pair.is_empty() { continue; }
        let (k, v) = if let Some(pos) = pair.find('=') {
            (&pair[..pos], &pair[pos+1..])
        } else {
            (pair, "")
        };
        let dk = decode_param(k);
        let dv = decode_param(v);
        keys.push(s(&dk));
        vals.push(s(&dv));
    }
    let mut o = Object::new();
    o.properties.insert("__keys".into(), arr_val(keys));
    o.properties.insert("__vals".into(), arr_val(vals));
    Value::Object(Arc::new(Mutex::new(o)))
}

fn decode_param(s: &str) -> String {
    let with_plus = s.replace('+', " ");
    percent_decode(&with_plus)
}

fn percent_decode(s: &str) -> String {
    let mut bytes = Vec::with_capacity(s.len());
    let src = s.as_bytes();
    let mut i = 0;
    while i < src.len() {
        if src[i] == b'%' && i + 2 < src.len() {
            if let Ok(hex) = std::str::from_utf8(&src[i+1..i+3]) {
                if let Ok(b) = u8::from_str_radix(hex, 16) {
                    bytes.push(b);
                    i += 3;
                    continue;
                }
            }
        }
        bytes.push(src[i]);
        i += 1;
    }
    String::from_utf8_lossy(&bytes).into_owned()
}

fn encode_param(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' | b'*' => {
                out.push(b as char);
            }
            b' ' => out.push('+'),
            _ => out.push_str(&format!("%{b:02X}")),
        }
    }
    out
}

fn get_keys(params: &Object) -> Vec<String> {
    if let Some(Value::Object(arr)) = params.properties.get("__keys") {
        let arr = arr.lock().unwrap();
        if let ObjectKind::Array(elems) = &arr.kind {
            return elems.iter().map(|v| match v {
                Value::String(s) => s.to_string(),
                _ => String::new(),
            }).collect();
        }
    }
    vec![]
}

fn get_vals(params: &Object) -> Vec<String> {
    if let Some(Value::Object(arr)) = params.properties.get("__vals") {
        let arr = arr.lock().unwrap();
        if let ObjectKind::Array(elems) = &arr.kind {
            return elems.iter().map(|v| match v {
                Value::String(s) => s.to_string(),
                _ => String::new(),
            }).collect();
        }
    }
    vec![]
}

fn set_kv(params: &mut Object, keys: Vec<String>, vals: Vec<String>) {
    params.properties.insert("__keys".into(), arr_val(keys.into_iter().map(|k| s(&k)).collect()));
    params.properties.insert("__vals".into(), arr_val(vals.into_iter().map(|v| s(&v)).collect()));
}

fn params_to_query(params: &Object) -> String {
    let keys = get_keys(params);
    let vals = get_vals(params);
    keys.iter().zip(vals.iter())
        .map(|(k, v)| format!("{}={}", encode_param(k), encode_param(v)))
        .collect::<Vec<_>>()
        .join("&")
}

fn get_sp_object(v: &Value) -> Option<std::sync::MutexGuard<Object>> {
    if let Value::Object(o) = v { Some(o.lock().unwrap()) } else { None }
}

pub fn register(vm: &mut VM) {
    // URL constructor
    vm.register_host_fn("node:url", "URL", Box::new(|_ctx, args| {
        let input = match args.first() { Some(Value::String(s)) => s.to_string(), _ => return Value::Undefined };
        let base = match args.get(1) { Some(Value::String(s)) => Some(s.to_string()), _ => None };
        match parse_url(&input, base.as_deref()) {
            Some(url) => build_url_obj(&url),
            None => Value::Undefined,
        }
    }));

    // URL.canParse
    vm.register_host_fn("node:url", "canParse", Box::new(|_ctx, args| {
        let input = match args.first() { Some(Value::String(s)) => s.to_string(), _ => return Value::Bool(false) };
        let base = match args.get(1) { Some(Value::String(s)) => Some(s.to_string()), _ => None };
        Value::Bool(parse_url(&input, base.as_deref()).is_some())
    }));

    // URLSearchParams constructor
    vm.register_host_fn("node:url", "URLSearchParams", Box::new(|_ctx, args| {
        let query = match args.first() {
            Some(Value::String(s)) => {
                let q = s.as_ref();
                if q.starts_with('?') { &q[1..] } else { q }.to_string()
            }
            _ => String::new(),
        };
        build_search_params(&query)
    }));

    // searchParamsGet
    vm.register_host_fn("node:url", "searchParamsGet", Box::new(|_ctx, args| {
        let key = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return Value::Null };
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let keys = get_keys(&sp);
            for (i, k) in keys.iter().enumerate() {
                if k == &key {
                    let vals = get_vals(&sp);
                    return s(&vals[i]);
                }
            }
        }
        Value::Null
    }));

    // searchParamsGetAll
    vm.register_host_fn("node:url", "searchParamsGetAll", Box::new(|_ctx, args| {
        let key = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return empty_array() };
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let keys = get_keys(&sp);
            let vals = get_vals(&sp);
            let matches: Vec<Value> = keys.iter().zip(vals.iter())
                .filter(|(k, _)| *k == &key)
                .map(|(_, v)| s(v))
                .collect();
            return arr_val(matches);
        }
        empty_array()
    }));

    // searchParamsHas
    vm.register_host_fn("node:url", "searchParamsHas", Box::new(|_ctx, args| {
        let key = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return Value::Bool(false) };
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            return Value::Bool(get_keys(&sp).contains(&key));
        }
        Value::Bool(false)
    }));

    // searchParamsSet
    vm.register_host_fn("node:url", "searchParamsSet", Box::new(|_ctx, args| {
        let key = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return Value::Undefined };
        let val = match args.get(2) { Some(Value::String(s)) => s.to_string(), _ => String::new() };
        if let Some(mut sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let mut keys = get_keys(&sp);
            let mut vals = get_vals(&sp);
            // Remove all existing entries with this key
            let mut i = 0;
            while i < keys.len() {
                if keys[i] == key { keys.remove(i); vals.remove(i); } else { i += 1; }
            }
            keys.push(key);
            vals.push(val);
            set_kv(&mut sp, keys, vals);
        }
        Value::Undefined
    }));

    // searchParamsAppend
    vm.register_host_fn("node:url", "searchParamsAppend", Box::new(|_ctx, args| {
        let key = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return Value::Undefined };
        let val = match args.get(2) { Some(Value::String(s)) => s.to_string(), _ => String::new() };
        if let Some(mut sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let mut keys = get_keys(&sp);
            let mut vals = get_vals(&sp);
            keys.push(key);
            vals.push(val);
            set_kv(&mut sp, keys, vals);
        }
        Value::Undefined
    }));

    // searchParamsDelete
    vm.register_host_fn("node:url", "searchParamsDelete", Box::new(|_ctx, args| {
        let key = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return Value::Undefined };
        if let Some(mut sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let mut keys = get_keys(&sp);
            let mut vals = get_vals(&sp);
            let mut i = 0;
            while i < keys.len() {
                if keys[i] == key { keys.remove(i); vals.remove(i); } else { i += 1; }
            }
            set_kv(&mut sp, keys, vals);
        }
        Value::Undefined
    }));

    // searchParamsToString
    vm.register_host_fn("node:url", "searchParamsToString", Box::new(|_ctx, args| {
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            return s(&params_to_query(&sp));
        }
        s("")
    }));

    // searchParamsSize
    vm.register_host_fn("node:url", "searchParamsSize", Box::new(|_ctx, args| {
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            return Value::I32(get_keys(&sp).len() as i32);
        }
        Value::I32(0)
    }));

    // searchParamsKeys
    vm.register_host_fn("node:url", "searchParamsKeys", Box::new(|_ctx, args| {
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let keys: Vec<Value> = get_keys(&sp).into_iter().map(|k| s(&k)).collect();
            return arr_val(keys);
        }
        empty_array()
    }));

    // searchParamsValues
    vm.register_host_fn("node:url", "searchParamsValues", Box::new(|_ctx, args| {
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let vals: Vec<Value> = get_vals(&sp).into_iter().map(|v| s(&v)).collect();
            return arr_val(vals);
        }
        empty_array()
    }));

    // searchParamsEntries
    vm.register_host_fn("node:url", "searchParamsEntries", Box::new(|_ctx, args| {
        if let Some(sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let keys = get_keys(&sp);
            let vals = get_vals(&sp);
            let pairs: Vec<Value> = keys.iter().zip(vals.iter())
                .map(|(k, v)| arr_val(vec![s(k), s(v)]))
                .collect();
            return arr_val(pairs);
        }
        empty_array()
    }));

    // searchParamsSort
    vm.register_host_fn("node:url", "searchParamsSort", Box::new(|_ctx, args| {
        if let Some(mut sp) = get_sp_object(args.first().unwrap_or(&Value::Undefined)) {
            let keys = get_keys(&sp);
            let vals = get_vals(&sp);
            let mut pairs: Vec<(String, String)> = keys.into_iter().zip(vals.into_iter()).collect();
            pairs.sort_by(|a, b| a.0.cmp(&b.0));
            let (ks, vs): (Vec<String>, Vec<String>) = pairs.into_iter().unzip();
            set_kv(&mut sp, ks, vs);
        }
        Value::Undefined
    }));

    // searchParamsForEach — stub
    vm.register_host_fn("node:url", "searchParamsForEach", Box::new(|_ctx, _args| Value::Undefined));

    // Legacy parse
    vm.register_host_fn("node:url", "parse", Box::new(|_ctx, args| {
        let input = match args.first() { Some(Value::String(s)) => s.to_string(), _ => return Value::Null };
        match parse_url(&input, None) {
            Some(url) => build_url_obj(&url),
            None => Value::Null,
        }
    }));

    // Legacy format
    vm.register_host_fn("node:url", "format", Box::new(|_ctx, args| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(href) = o.properties.get("href") {
                return href.clone();
            }
        }
        s("")
    }));

    // Legacy resolve
    vm.register_host_fn("node:url", "resolve", Box::new(|_ctx, args| {
        let from = match args.first() { Some(Value::String(s)) => s.to_string(), _ => return s("") };
        let to = match args.get(1) { Some(Value::String(s)) => s.to_string(), _ => return s("") };
        match parse_url(&to, Some(&from)) {
            Some(url) => s(url.as_str()),
            None => s(""),
        }
    }));

    // fileURLToPath
    vm.register_host_fn("node:url", "fileURLToPath", Box::new(|_ctx, args| {
        let input = match args.first() { Some(Value::String(s)) => s.to_string(), _ => return s("") };
        if let Some(url) = parse_url(&input, None) {
            if url.scheme() == "file" {
                let path = url.path().to_string();
                // Decode percent-encoding in path
                return s(percent_decode(&path));
            }
        }
        // Fallback: strip "file://"
        if let Some(rest) = input.strip_prefix("file://") {
            return s(rest);
        }
        s(&input)
    }));

    // pathToFileURL
    vm.register_host_fn("node:url", "pathToFileURL", Box::new(|_ctx, args| {
        let path = match args.first() { Some(Value::String(s)) => s.to_string(), _ => return Value::Undefined };
        let file_url = format!("file://{path}");
        match parse_url(&file_url, None) {
            Some(url) => build_url_obj(&url),
            None => {
                // Manually build minimal URL object
                let mut o = Object::new();
                o.properties.insert("href".into(), s(&file_url));
                o.properties.insert("protocol".into(), s("file:"));
                o.properties.insert("pathname".into(), s(&path));
                Value::Object(Arc::new(Mutex::new(o)))
            }
        }
    }));

    // domainToASCII / domainToUnicode — pass-through for ASCII domains
    vm.register_host_fn("node:url", "domainToASCII", Box::new(|_ctx, args| {
        match args.first() { Some(Value::String(s)) => Value::String(s.clone()), _ => s("") }
    }));
    vm.register_host_fn("node:url", "domainToUnicode", Box::new(|_ctx, args| {
        match args.first() { Some(Value::String(s)) => Value::String(s.clone()), _ => s("") }
    }));

    // urlToString / urlToJSON
    vm.register_host_fn("node:url", "urlToString", Box::new(|_ctx, args| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(href) = o.properties.get("href") { return href.clone(); }
        }
        s("")
    }));
    vm.register_host_fn("node:url", "urlToJSON", Box::new(|_ctx, args| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            if let Some(href) = o.properties.get("href") { return href.clone(); }
        }
        s("")
    }));
}
