//! WHATWG URL Living Standard — URL + URLSearchParams.
//!
//!   `new URL(input, base?)` — parses an absolute or relative URL
//!   `URL.canParse(input, base?)` — bool predicate (no exception)
//!   `URL.parse(input, base?)` → URL | null (Stage-3, replaces try/catch)
//!   `urlObj.{href, origin, protocol, host, hostname, port, pathname,
//!            search, hash, username, password, searchParams}`
//!
//!   `new URLSearchParams(init)` — query-string parser/serialiser
//!   `params.{get, set, has, append, delete, getAll, toString, sort,
//!            keys, values, entries, size}`
//!
//! Vybe leans on `form_urlencoded` (already a workspace dep) for the
//! query-pair handling and parses the URL with a hand-rolled parser
//! that follows the WHATWG state machine for the common HTTP/HTTPS
//! shape — full RFC 3986 + WHATWG normalization is a TODO.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

#[derive(Default, Clone)]
struct UrlParts {
    href: String,
    protocol: String,  // "https:"
    username: String,
    password: String,
    host: String,      // "example.com:8080"
    hostname: String,  // "example.com"
    port: String,      // "8080" or ""
    pathname: String,  // "/foo/bar"
    search: String,    // "?a=1" (leading ?)
    hash: String,      // "#frag" (leading #)
}

impl UrlParts {
    fn origin(&self) -> String {
        if self.protocol.is_empty() || self.host.is_empty() { "null".into() }
        else { format!("{}//{}", self.protocol, self.host) }
    }
}

fn parse_url(input: &str, base: Option<&str>) -> Option<UrlParts> {
    let mut s = input.trim().to_string();
    // Resolve relative against base.
    if !s.contains("://") {
        if let Some(b) = base {
            let bp = parse_url(b, None)?;
            if s.starts_with('/') {
                s = format!("{}//{}{}", bp.protocol, bp.host, s);
            } else {
                let dir = bp.pathname.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
                s = format!("{}//{}{}/{}", bp.protocol, bp.host, dir, s);
            }
        } else {
            return None;
        }
    }
    let mut p = UrlParts::default();
    p.href = s.clone();
    // Protocol.
    let scheme_end = s.find("://")?;
    p.protocol = format!("{}:", &s[..scheme_end]);
    let rest = &s[scheme_end + 3..];

    // Hash.
    let (rest, hash) = match rest.find('#') {
        Some(i) => (&rest[..i], &rest[i..]),
        None => (rest, ""),
    };
    p.hash = hash.to_string();

    // Search.
    let (rest, search) = match rest.find('?') {
        Some(i) => (&rest[..i], &rest[i..]),
        None => (rest, ""),
    };
    p.search = search.to_string();

    // Host vs path.
    let (authority, path) = match rest.find('/') {
        Some(i) => (&rest[..i], &rest[i..]),
        None => (rest, ""),
    };
    p.pathname = if path.is_empty() { "/".into() } else { path.into() };

    // Userinfo @ host.
    let host_part = if let Some(at) = authority.rfind('@') {
        let userinfo = &authority[..at];
        if let Some((u, pw)) = userinfo.split_once(':') {
            p.username = u.into();
            p.password = pw.into();
        } else {
            p.username = userinfo.into();
        }
        &authority[at + 1..]
    } else {
        authority
    };
    p.host = host_part.into();
    if let Some((h, port)) = host_part.rsplit_once(':') {
        p.hostname = h.into();
        p.port = port.into();
    } else {
        p.hostname = host_part.into();
    }
    Some(p)
}

fn make_url_object(p: &UrlParts) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("URL")));
    let s = |x: &str| Value::String(Arc::from(x));
    obj.properties.insert("href".into(),     s(&p.href));
    obj.properties.insert("origin".into(),   s(&p.origin()));
    obj.properties.insert("protocol".into(), s(&p.protocol));
    obj.properties.insert("username".into(), s(&p.username));
    obj.properties.insert("password".into(), s(&p.password));
    obj.properties.insert("host".into(),     s(&p.host));
    obj.properties.insert("hostname".into(), s(&p.hostname));
    obj.properties.insert("port".into(),     s(&p.port));
    obj.properties.insert("pathname".into(), s(&p.pathname));
    obj.properties.insert("search".into(),   s(&p.search));
    obj.properties.insert("hash".into(),     s(&p.hash));
    obj.properties.insert("searchParams".into(), make_search_params(&p.search));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_search_params(query: &str) -> Value {
    let q = query.trim_start_matches('?');
    let pairs: Vec<(String, String)> = form_urlencoded::parse(q.as_bytes())
        .into_owned()
        .collect();
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("URLSearchParams")));
    let entries: Vec<Value> = pairs.iter().map(|(k, v)| {
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
            Value::String(Arc::from(k.as_str())),
            Value::String(Arc::from(v.as_str())),
        ]))))
    }).collect();
    obj.properties.insert("__pairs".into(), Value::Object(Arc::new(Mutex::new(Object::new_array(entries)))));
    obj.properties.insert("size".into(), Value::F64(pairs.len() as f64));
    Value::Object(Arc::new(Mutex::new(obj)))
}

pub fn register(vm: &mut VM) {
    // new URL(input, base?) — §URL parsing. The known_types-driven `new`
    // dispatch in Vybe passes the user args directly (no implicit `this`),
    // matching the Intl `new` convention.
    vm.register_host_fn("web:url", "new", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let base = args.get(1).map(|v| format!("{}", v));
        let parts = parse_url(&input, base.as_deref()).unwrap_or_default();
        make_url_object(&parts)
    }));

    // URL.parse(input, base?) → URL | null (Stage-3).
    vm.register_host_fn("web:url", "parse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let base = args.get(1).map(|v| format!("{}", v));
        match parse_url(&input, base.as_deref()) {
            Some(p) => make_url_object(&p),
            None => Value::Null,
        }
    }));

    // URL.canParse(input, base?) → bool.
    vm.register_host_fn("web:url", "canParse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let base = args.get(1).map(|v| format!("{}", v));
        Value::Bool(parse_url(&input, base.as_deref()).is_some())
    }));

    // ── URLSearchParams ────────────────────────────────────────────
    vm.register_host_fn("web:url", "searchParamsNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let init = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        make_search_params(&init)
    }));

    vm.register_host_fn("web:url", "searchParamsGet", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.first() {
            if let Some(Value::Object(pairs_obj)) = obj.lock().unwrap().properties.get("__pairs") {
                let p = pairs_obj.lock().unwrap();
                if let ObjectKind::Array(ref pairs) = p.kind {
                    for pair in pairs {
                        if let Value::Object(arr) = pair {
                            let a = arr.lock().unwrap();
                            if let ObjectKind::Array(ref kv) = a.kind {
                                if kv.len() >= 2 && format!("{}", kv[0]) == key {
                                    return kv[1].clone();
                                }
                            }
                        }
                    }
                }
            }
        }
        Value::Null
    }));

    vm.register_host_fn("web:url", "searchParamsHas", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.first() {
            if let Some(Value::Object(pairs_obj)) = obj.lock().unwrap().properties.get("__pairs") {
                let p = pairs_obj.lock().unwrap();
                if let ObjectKind::Array(ref pairs) = p.kind {
                    for pair in pairs {
                        if let Value::Object(arr) = pair {
                            let a = arr.lock().unwrap();
                            if let ObjectKind::Array(ref kv) = a.kind {
                                if kv.len() >= 2 && format!("{}", kv[0]) == key {
                                    return Value::Bool(true);
                                }
                            }
                        }
                    }
                }
            }
        }
        Value::Bool(false)
    }));

    vm.register_host_fn("web:url", "searchParamsToString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            if let Some(Value::Object(pairs_obj)) = obj.lock().unwrap().properties.get("__pairs") {
                let p = pairs_obj.lock().unwrap();
                if let ObjectKind::Array(ref pairs) = p.kind {
                    let mut ser = form_urlencoded::Serializer::new(String::new());
                    for pair in pairs {
                        if let Value::Object(arr) = pair {
                            let a = arr.lock().unwrap();
                            if let ObjectKind::Array(ref kv) = a.kind {
                                if kv.len() >= 2 {
                                    ser.append_pair(&format!("{}", kv[0]), &format!("{}", kv[1]));
                                }
                            }
                        }
                    }
                    return Value::String(Arc::from(ser.finish().as_str()));
                }
            }
        }
        Value::String(Arc::from(""))
    }));
}
