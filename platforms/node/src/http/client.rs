//! Node.js `http` client-side API surface.
//!
//! Registers STATUS_CODES, METHODS, Agent, createServer, request, get,
//! IncomingMessage, ServerResponse and companion constants.

use std::sync::Arc;
use vybe_runtime::VM;
use vybe_runtime::value::{Object, Value};

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn stub() -> Value {
    s("")
}

fn make_empty_obj() -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new()))
}

fn make_empty_arr() -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())))
}

fn make_client_request(method: &str, host: &str, path: &str) -> Value {
    let mut o = Object::new();
    o.properties.insert("method".into(), s(method));
    o.properties.insert("path".into(), s(path));
    o.properties.insert("host".into(), s(host));
    o.properties.insert("protocol".into(), s("http:"));
    o.properties.insert("finished".into(), Value::Bool(false));
    o.properties.insert("destroyed".into(), Value::Bool(false));
    for m in [
        "end",
        "write",
        "setHeader",
        "getHeader",
        "removeHeader",
        "destroy",
        "setTimeout",
        "abort",
        "on",
        "once",
        "emit",
        "pipe",
    ] {
        o.properties.insert(m.into(), stub());
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn make_agent(max_sockets: i32) -> Value {
    let mut o = Object::new();
    o.properties
        .insert("maxSockets".into(), Value::I32(max_sockets));
    o.properties
        .insert("maxFreeSockets".into(), Value::I32(256));
    o.properties
        .insert("maxTotalSockets".into(), Value::F64(f64::INFINITY));
    o.properties.insert("sockets".into(), make_empty_obj());
    o.properties.insert("freeSockets".into(), make_empty_obj());
    o.properties.insert("requests".into(), make_empty_obj());
    o.properties.insert("keepAlive".into(), Value::Bool(false));
    for m in [
        "destroy",
        "getName",
        "createConnection",
        "addRequest",
        "reuseSocket",
        "removeSocket",
    ] {
        o.properties.insert(m.into(), stub());
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn parse_url(url: &str) -> (String, String) {
    let url = url
        .trim_start_matches("http://")
        .trim_start_matches("https://");
    if let Some(slash_pos) = url.find('/') {
        (url[..slash_pos].to_string(), url[slash_pos..].to_string())
    } else {
        (url.to_string(), "/".to_string())
    }
}

fn get_str_prop(args: &[Value], idx: usize, key: &str, default: &str) -> String {
    match args.get(idx) {
        Some(Value::Object(opts)) => {
            let o = opts.lock().unwrap();
            match o.properties.get(key) {
                Some(Value::String(s)) => s.to_string(),
                _ => default.to_string() }
        }
        _ => default.to_string() }
}

pub fn register(vm: &mut VM) {
    // STATUS_CODES
    vm.register_host_fn(
        "node:http",
        "STATUS_CODES",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            let codes: &[(&str, &str)] = &[
                ("100", "Continue"),
                ("101", "Switching Protocols"),
                ("102", "Processing"),
                ("200", "OK"),
                ("201", "Created"),
                ("202", "Accepted"),
                ("203", "Non-Authoritative Information"),
                ("204", "No Content"),
                ("205", "Reset Content"),
                ("206", "Partial Content"),
                ("207", "Multi-Status"),
                ("208", "Already Reported"),
                ("226", "IM Used"),
                ("300", "Multiple Choices"),
                ("301", "Moved Permanently"),
                ("302", "Found"),
                ("303", "See Other"),
                ("304", "Not Modified"),
                ("305", "Use Proxy"),
                ("307", "Temporary Redirect"),
                ("308", "Permanent Redirect"),
                ("400", "Bad Request"),
                ("401", "Unauthorized"),
                ("402", "Payment Required"),
                ("403", "Forbidden"),
                ("404", "Not Found"),
                ("405", "Method Not Allowed"),
                ("406", "Not Acceptable"),
                ("407", "Proxy Authentication Required"),
                ("408", "Request Timeout"),
                ("409", "Conflict"),
                ("410", "Gone"),
                ("411", "Length Required"),
                ("412", "Precondition Failed"),
                ("413", "Payload Too Large"),
                ("414", "URI Too Long"),
                ("415", "Unsupported Media Type"),
                ("416", "Range Not Satisfiable"),
                ("417", "Expectation Failed"),
                ("418", "I'm a Teapot"),
                ("421", "Misdirected Request"),
                ("422", "Unprocessable Entity"),
                ("423", "Locked"),
                ("424", "Failed Dependency"),
                ("425", "Too Early"),
                ("426", "Upgrade Required"),
                ("428", "Precondition Required"),
                ("429", "Too Many Requests"),
                ("431", "Request Header Fields Too Large"),
                ("451", "Unavailable For Legal Reasons"),
                ("500", "Internal Server Error"),
                ("501", "Not Implemented"),
                ("502", "Bad Gateway"),
                ("503", "Service Unavailable"),
                ("504", "Gateway Timeout"),
                ("505", "HTTP Version Not Supported"),
                ("506", "Variant Also Negotiates"),
                ("507", "Insufficient Storage"),
                ("508", "Loop Detected"),
                ("510", "Not Extended"),
                ("511", "Network Authentication Required"),
            ];
            for (code, text) in codes {
                o.properties
                    .insert((*code).into(), Value::String(Arc::from(*text)));
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // METHODS
    vm.register_host_fn(
        "node:http",
        "METHODS",
        Box::new(|_ctx, _args| {
            let methods = [
                "ACL",
                "BIND",
                "CHECKOUT",
                "CONNECT",
                "COPY",
                "DELETE",
                "GET",
                "HEAD",
                "LINK",
                "LOCK",
                "M-SEARCH",
                "MERGE",
                "MKACTIVITY",
                "MKCALENDAR",
                "MKCOL",
                "MOVE",
                "NOTIFY",
                "OPTIONS",
                "PATCH",
                "POST",
                "PROPFIND",
                "PROPPATCH",
                "PURGE",
                "PUT",
                "REBIND",
                "REPORT",
                "SEARCH",
                "SOURCE",
                "SUBSCRIBE",
                "TRACE",
                "UNBIND",
                "UNLINK",
                "UNLOCK",
                "UNSUBSCRIBE",
            ];
            let elems: Vec<Value> = methods
                .iter()
                .map(|m| Value::String(Arc::from(*m)))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
        }),
    );

    // maxHeaderSize
    vm.register_host_fn(
        "node:http",
        "maxHeaderSize",
        Box::new(|_ctx, _args| Value::I32(16384)),
    );

    // globalAgent
    vm.register_host_fn(
        "node:http",
        "globalAgent",
        Box::new(|_ctx, _args| make_agent(256)),
    );

    // validateHeaderName
    vm.register_host_fn(
        "node:http",
        "validateHeaderName",
        Box::new(|_ctx, args| {
            let name = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::Bool(false) };
            // Header names must only contain token characters (no space, no control chars)
            let valid = name
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || "-_!#$%&'*+.^`|~".contains(c));
            if valid {
                Value::Undefined
            } else {
                Value::Bool(false)
            }
        }),
    );

    // validateHeaderValue
    vm.register_host_fn(
        "node:http",
        "validateHeaderValue",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    // setMaxIdleHTTPParsers
    vm.register_host_fn(
        "node:http",
        "setMaxIdleHTTPParsers",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    // createServer([opts, cb]) → Server
    vm.register_host_fn(
        "node:http",
        "createServer",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("listening".into(), Value::Bool(false));
            o.properties.insert("timeout".into(), Value::I32(0));
            o.properties
                .insert("keepAliveTimeout".into(), Value::I32(5000));
            o.properties
                .insert("headersTimeout".into(), Value::I32(60000));
            o.properties
                .insert("requestTimeout".into(), Value::I32(300000));
            o.properties
                .insert("maxConnections".into(), Value::Undefined);
            for m in [
                "listen",
                "close",
                "address",
                "setTimeout",
                "getConnections",
                "on",
                "once",
                "emit",
                "ref",
                "unref",
            ] {
                o.properties.insert(m.into(), stub());
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // request(options[, cb]) → ClientRequest
    vm.register_host_fn(
        "node:http",
        "request",
        Box::new(|_ctx, args| {
            let host = get_str_prop(args, 0, "host", "localhost");
            let path = get_str_prop(args, 0, "path", "/");
            let method = get_str_prop(args, 0, "method", "GET");
            make_client_request(&method, &host, &path)
        }),
    );

    // get(url[, opts, cb]) → ClientRequest (GET)
    vm.register_host_fn(
        "node:http",
        "get",
        Box::new(|_ctx, args| {
            let (host, path) = match args.first() {
                Some(Value::String(url)) => parse_url(url),
                _ => {
                    let host = get_str_prop(args, 0, "host", "localhost");
                    let path = get_str_prop(args, 0, "path", "/");
                    (host, path)
                }
            };
            make_client_request("GET", &host, &path)
        }),
    );

    // Agent([opts]) → Agent
    vm.register_host_fn(
        "node:http",
        "Agent",
        Box::new(|_ctx, args| {
            let max_sockets = match args.first() {
                Some(Value::Object(opts)) => {
                    let o = opts.lock().unwrap();
                    match o.properties.get("maxSockets") {
                        Some(Value::I32(n)) => *n,
                        Some(Value::F64(f)) => *f as i32,
                        _ => 256 }
                }
                _ => 256 };
            make_agent(max_sockets)
        }),
    );

    // IncomingMessage([socket]) → IncomingMessage object
    vm.register_host_fn(
        "node:http",
        "IncomingMessage",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("method".into(), Value::Null);
            o.properties.insert("url".into(), Value::Null);
            o.properties.insert("headers".into(), make_empty_obj());
            o.properties.insert("rawHeaders".into(), make_empty_arr());
            o.properties.insert("httpVersion".into(), s("1.1"));
            o.properties
                .insert("httpVersionMajor".into(), Value::I32(1));
            o.properties
                .insert("httpVersionMinor".into(), Value::I32(1));
            o.properties.insert("statusCode".into(), Value::Null);
            o.properties.insert("statusMessage".into(), Value::Null);
            o.properties.insert("complete".into(), Value::Bool(false));
            o.properties.insert("trailers".into(), make_empty_obj());
            o.properties.insert("rawTrailers".into(), make_empty_arr());
            o.properties.insert("readable".into(), Value::Bool(true));
            o.properties.insert("destroyed".into(), Value::Bool(false));
            for m in [
                "on",
                "once",
                "emit",
                "pipe",
                "pause",
                "resume",
                "destroy",
                "read",
                "setTimeout",
            ] {
                o.properties.insert(m.into(), stub());
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ServerResponse([req]) → ServerResponse object
    vm.register_host_fn(
        "node:http",
        "ServerResponse",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("statusCode".into(), Value::I32(200));
            o.properties.insert("statusMessage".into(), s("OK"));
            o.properties
                .insert("headersSent".into(), Value::Bool(false));
            o.properties
                .insert("writableEnded".into(), Value::Bool(false));
            o.properties
                .insert("writableFinished".into(), Value::Bool(false));
            o.properties.insert("finished".into(), Value::Bool(false));
            o.properties.insert("writable".into(), Value::Bool(true));
            o.properties
                .insert("chunkedEncoding".into(), Value::Bool(false));
            for m in [
                "writeHead",
                "setHeader",
                "getHeader",
                "getHeaders",
                "getHeaderNames",
                "hasHeader",
                "removeHeader",
                "write",
                "end",
                "flushHeaders",
                "on",
                "once",
                "emit",
                "setTimeout",
                "addTrailers",
            ] {
                o.properties.insert(m.into(), stub());
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );
}
