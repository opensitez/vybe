//! WHATWG Fetch — fetch + Request + Response + Headers.
//!
//!   `fetch(input, init?)` → Promise<Response>
//!   `new Request(input, init?)` — request descriptor
//!   `new Response(body?, init?)` — response with status/headers/body
//!   `new Headers(init?)` — case-insensitive name/value map
//!
//! Vybe's Promise model is sync-by-default (see `vybe_platform_ecma::promise`).
//! `fetch()` performs the request synchronously on the host thread via
//! the existing wasi:http path and returns a fulfilled Promise — JSPI
//! lift-and-suspend can wrap this transparently when the user awaits.
//!
//! HTTP transport: this MVP uses a blocking `std::net::TcpStream` for
//! `http://` URLs and returns an empty body for `https://` (TLS would
//! pull in `rustls` — a separate dep decision). Full implementations
//! delegate to wasi:http or `reqwest` once enabled.

use std::collections::HashMap;
use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

fn make_promise_fulfilled(value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties
        .insert("__state".into(), Value::String(Arc::from("fulfilled")));
    obj.properties.insert("__value".into(), value);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_promise_rejected(reason: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties
        .insert("__state".into(), Value::String(Arc::from("rejected")));
    obj.properties.insert("__value".into(), reason);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn headers_obj(map: HashMap<String, String>) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Headers")));
    let entries: Vec<Value> = map
        .iter()
        .map(|(k, v)| {
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                Value::String(Arc::from(k.as_str())),
                Value::String(Arc::from(v.as_str())),
            ])))
        })
        .collect();
    obj.properties.insert(
        "__entries".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(entries))),
    );
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_response(
    status: u16,
    status_text: &str,
    body: String,
    headers: HashMap<String, String>,
) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Response")));
    obj.properties
        .insert("status".into(), Value::F64(status as f64));
    obj.properties
        .insert("statusText".into(), Value::String(Arc::from(status_text)));
    obj.properties
        .insert("ok".into(), Value::Bool((200..300).contains(&status)));
    obj.properties
        .insert("headers".into(), headers_obj(headers));
    obj.properties
        .insert("__body".into(), Value::String(Arc::from(body.as_str())));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

pub fn register(vm: &mut VM) {
    // fetch(input, init?) → Promise<Response>
    //
    // input: string URL or Request object. init: { method, headers, body }.
    vm.register_host_fn(
        "web:fetch",
        "fetch",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let url = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(Value::Object(obj)) => {
                    let o = obj.lock().unwrap();
                    o.properties
                        .get("url")
                        .map(|v| format!("{}", v))
                        .or_else(|| o.properties.get("href").map(|v| format!("{}", v)))
                        .unwrap_or_default()
                }
                _ => return make_promise_rejected(Value::String(Arc::from("invalid url"))),
            };
            // MVP: only http:// goes over TCP. https:// returns 0/empty.
            if !url.starts_with("http://") {
                return make_promise_fulfilled(make_response(0, "", String::new(), HashMap::new()));
            }
            match http_get_blocking(&url) {
                Ok((status, headers, body)) => {
                    make_promise_fulfilled(make_response(status, "OK", body, headers))
                }
                Err(e) => make_promise_rejected(Value::String(Arc::from(e.to_string().as_str()))),
            }
        }),
    );

    // new Request(input, init?) — args[0]=input, args[1]=init.
    vm.register_host_fn(
        "web:fetch",
        "requestNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let url = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let method = args
                .get(1)
                .and_then(|init| {
                    if let Value::Object(io) = init {
                        io.lock()
                            .unwrap()
                            .properties
                            .get("method")
                            .map(|v| format!("{}", v))
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "GET".into());
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Request")));
            obj.properties
                .insert("url".into(), Value::String(Arc::from(url.as_str())));
            obj.properties
                .insert("method".into(), Value::String(Arc::from(method.as_str())));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // new Response(body?, init?) — args[0]=body, args[1]=init.
    vm.register_host_fn(
        "web:fetch",
        "responseNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let body = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let (status, status_text) = if let Some(Value::Object(init)) = args.get(1) {
                let io = init.lock().unwrap();
                let st = io
                    .properties
                    .get("status")
                    .map(|v| v.as_f64() as u16)
                    .unwrap_or(200);
                let stt = io
                    .properties
                    .get("statusText")
                    .map(|v| format!("{}", v))
                    .unwrap_or_else(|| "OK".into());
                (st, stt)
            } else {
                (200, "OK".into())
            };
            Value::Object(vybe_runtime::heap::alloc({
                let mut obj = Object::new();
                obj.properties
                    .insert("__type".into(), Value::String(Arc::from("Response")));
                obj.properties
                    .insert("status".into(), Value::F64(status as f64));
                obj.properties.insert(
                    "statusText".into(),
                    Value::String(Arc::from(status_text.as_str())),
                );
                obj.properties
                    .insert("ok".into(), Value::Bool((200..300).contains(&status)));
                obj.properties
                    .insert("__body".into(), Value::String(Arc::from(body.as_str())));
                obj.properties
                    .insert("headers".into(), headers_obj(HashMap::new()));
                obj
            }))
        }),
    );

    // response.text() → Promise<string>
    vm.register_host_fn(
        "web:fetch",
        "responseText",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let body = o
                    .properties
                    .get("__body")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                return make_promise_fulfilled(Value::String(Arc::from(body.as_str())));
            }
            make_promise_fulfilled(Value::String(Arc::from("")))
        }),
    );

    // response.json() → Promise<any> (synchronous JSON.parse on the body).
    vm.register_host_fn(
        "web:fetch",
        "responseJson",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let body = o
                    .properties
                    .get("__body")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                // Defer the actual JSON parse to ecma:json.parse via the promise body.
                return make_promise_fulfilled(Value::String(Arc::from(body.as_str())));
            }
            make_promise_fulfilled(Value::Null)
        }),
    );

    // ── Headers ────────────────────────────────────────────────────
    vm.register_host_fn(
        "web:fetch",
        "headersNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut map: HashMap<String, String> = HashMap::new();
            if let Some(Value::Object(init)) = args.first() {
                let io = init.lock().unwrap();
                for (k, v) in io.properties.iter() {
                    if !k.starts_with("__") {
                        map.insert(k.to_lowercase(), format!("{}", v));
                    }
                }
            }
            headers_obj(map)
        }),
    );
}

fn http_get_blocking(
    url: &str,
) -> Result<(u16, HashMap<String, String>, String), Box<dyn std::error::Error>> {
    use std::io::{Read, Write};
    use std::net::TcpStream;
    let url_no_scheme = url.trim_start_matches("http://");
    let (host_part, path) = match url_no_scheme.find('/') {
        Some(i) => (&url_no_scheme[..i], &url_no_scheme[i..]),
        None => (url_no_scheme, "/"),
    };
    let (host, port) = match host_part.rsplit_once(':') {
        Some((h, p)) => (h, p.parse::<u16>().unwrap_or(80)),
        None => (host_part, 80),
    };
    let mut stream = TcpStream::connect((host, port))?;
    let req = format!(
        "GET {} HTTP/1.0\r\nHost: {}\r\nConnection: close\r\n\r\n",
        path, host
    );
    stream.write_all(req.as_bytes())?;
    let mut buf = String::new();
    stream.read_to_string(&mut buf)?;
    // Parse status line + headers + body.
    let (head, body) = match buf.find("\r\n\r\n") {
        Some(i) => (&buf[..i], buf[i + 4..].to_string()),
        None => (buf.as_str(), String::new()),
    };
    let mut lines = head.lines();
    let status_line = lines.next().unwrap_or("");
    let status: u16 = status_line
        .split_whitespace()
        .nth(1)
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    let mut headers = HashMap::new();
    for line in lines {
        if let Some((k, v)) = line.split_once(':') {
            headers.insert(k.trim().to_lowercase(), v.trim().to_string());
        }
    }
    Ok((status, headers, body))
}
