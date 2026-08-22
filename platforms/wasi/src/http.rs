//! `wasi:http@0.3.1` — `types`, `client` and `handler`.
//!
//! The client half is on the 0.3.1 model: build a `request`
//! (`[static]request.new` + the `set-*` methods), hand it to `client.send`,
//! and read the response through `response.get-status-code` /
//! `response.get-headers`. `handler.handle` shares the implementation, which
//! the spec explicitly allows — a `client.send` import may be linked directly
//! to a `handler.handle` export.
//!
//! NOT yet on 0.3.1: the SERVER half and the 0.2 resource pairs it rests on.
//! 0.3.1 declares four resources — `fields`, `request`, `response`,
//! `request-options` — where this file still also registers `incoming-request`,
//! `outgoing-response`, `incoming-response`, `response-outparam`,
//! `outgoing-body`, `incoming-body`, `future-trailers` and
//! `future-incoming-response`. 0.3 collapsed the incoming/outgoing pairs into
//! one resource each and replaced `response-outparam` with `handle` simply
//! RETURNING the response. `interface_coverage.rs` reports each of those by
//! name; they come out with the server rewrite, not before, because
//! `http_incoming_server.rs` is built on them.

use std::collections::HashMap;
use std::io::{Read, Write};
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::{Arc, Mutex};

use vybe_runtime::typedef::TypeDef;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

const KIND_HEADERS: &str = "headers";
const KIND_OUTGOING_REQUEST: &str = "outgoing-request";
const KIND_REQUEST_OPTIONS: &str = "request-options";
const KIND_INCOMING_RESPONSE: &str = "incoming-response";
const KIND_INCOMING_BODY: &str = "incoming-body";
const KIND_OUTGOING_RESPONSE: &str = "outgoing-response";
const KIND_OUTGOING_BODY: &str = "outgoing-body";
const KIND_INCOMING_REQUEST: &str = "incoming-request";
const KIND_RESPONSE_OUTPARAM: &str = "response-outparam";

#[derive(Clone, Copy)]
struct HttpTypeIds {
    headers: usize,
    outgoing_request: usize,
    request_options: usize,
    incoming_response: usize,
    future_incoming_response: usize,
    incoming_body: usize,
    future_trailers: usize,
    outgoing_response: usize,
    outgoing_body: usize,
    _incoming_request: usize,
    _response_outparam: usize,
}

#[derive(Debug, Clone)]
struct HeadersResource {
    entries: Vec<(String, Vec<u8>)>,
}

#[derive(Debug, Clone)]
struct OutgoingRequestResource {
    headers_id: u32,
    /// 0.3 `request.new(headers, contents, trailers, options)` — `options` is
    /// `option<request-options>`, surfaced again by `request.get-options`.
    options_id: Option<u32>,
    /// The `contents` stream, drained at construction. 0.2 reached a request
    /// body through `outgoing-request.body` -> `outgoing-body.write`; 0.3.1
    /// deleted both, so the stream handed to `new` is the only body there is.
    body: Vec<u8>,
    method: String,
    path_with_query: Option<String>,
    scheme: Option<String>,
    authority: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct RequestOptionsResource {
    connect_timeout_ns: Option<u64>,
    first_byte_timeout_ns: Option<u64>,
    between_bytes_timeout_ns: Option<u64>,
}

#[derive(Debug, Clone)]
struct IncomingBodyResource {
    body: Vec<u8>,
    #[allow(dead_code)]
    position: usize,
}

#[derive(Debug, Clone)]
struct OutgoingResponseResource {
    status: u16,
    headers_id: u32,
    body_id: Option<u32>,
}

#[derive(Debug, Clone)]
struct OutgoingBodyResource {
    bytes: Vec<u8>,
    /// §outgoing-body.finish must be called exactly once.
    finished: bool,
    /// Optional trailers supplied to `finish`.
    trailers: Vec<(String, Vec<u8>)>,
}

/// A server-side request. Built by the host from whatever transport it is
/// serving (hyper, in `vybex --serve`), then handed to guest code as the
/// `incoming-request` resource of `wasi:http/incoming-handler`.
#[derive(Debug, Clone)]
struct IncomingRequestResource {
    method: String,
    path_with_query: Option<String>,
    scheme: Option<String>,
    authority: Option<String>,
    headers_id: u32,
    body: Vec<u8>,
    /// §incoming-request.consume: "Will only return success at most once, and
    /// subsequent calls will return error."
    consumed: bool,
}

/// The write end of a server response. `response-outparam.set` stores the
/// guest's `result<outgoing-response, error-code>` here; the host reads it back
/// after the handler returns.
#[derive(Debug, Clone, Default)]
struct ResponseOutparamResource {
    response_id: Option<u32>,
    error: Option<String>,
    informational: Vec<(u16, u32)>,
    set: bool,
}

#[derive(Debug, Clone)]
struct IncomingResponseResource {
    status: u16,
    headers_id: u32,
    body: String,
}


#[derive(Default)]
struct Registry {
    headers: HashMap<u32, HeadersResource>,
    outgoing_requests: HashMap<u32, OutgoingRequestResource>,
    request_options: HashMap<u32, RequestOptionsResource>,
    incoming_responses: HashMap<u32, IncomingResponseResource>,
    incoming_bodies: HashMap<u32, IncomingBodyResource>,
    outgoing_responses: HashMap<u32, OutgoingResponseResource>,
    outgoing_bodies: HashMap<u32, OutgoingBodyResource>,
    incoming_requests: HashMap<u32, IncomingRequestResource>,
    response_outparams: HashMap<u32, ResponseOutparamResource>,
}

/// Resource ids, OUTSIDE the registry so clearing tenant data cannot rewind
/// them and reissue a handle another tenant still holds.
static NEXT_ID: AtomicU32 = AtomicU32::new(1);

/// A free function, not a `Registry` method: the counter is not in the
/// registry, and a `alloc_id()` would read as though it were.
fn alloc_id() -> u32 {
    NEXT_ID.fetch_add(1, Ordering::Relaxed)
}

/// Every HTTP resource this program has open — headers, in-flight requests,
/// responses and bodies.
///
/// All of it is per-program, and VM-owned ([`vybe_runtime::resources`]) so the
/// VM drops it on `reset_to` without this module taking part. As a static it
/// let the next program in a reused VM read another tenant's response bodies
/// and headers through a handle it never created.
fn registry() -> &'static Mutex<Registry> {
    vybe_runtime::resources::get::<Registry>()
}

fn make_resource(kind: &str, id: u32, type_id: usize) -> Value {
    let mut object = Object::new();
    object.type_id = type_id;
    object
        .properties
        .insert("__wasi_kind".into(), Value::String(Arc::from(kind)));
    object
        .properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn resource_id(value: &Value, expected_kind: &str) -> Option<u32> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        let kind_ok = matches!(
            object.properties.get("__wasi_kind"),
            Some(Value::String(kind)) if kind.as_ref() == expected_kind
        );
        if !kind_ok {
            return None;
        }
        if let Some(Value::F64(id)) = object.properties.get("__wasi_id") {
            return Some(*id as u32);
        }
    }
    None
}

fn err(code: &str) -> Value {
    let mut object = Object::new();
    object
        .properties
        .insert("__wasi_error".into(), Value::String(Arc::from(code)));
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn string_arg(args: &[Value], idx: usize) -> Option<String> {
    match args.get(idx) {
        Some(Value::String(text)) => Some(text.to_string()),
        Some(Value::Null) | None => None,
        Some(value) => Some(format!("{}", value)),
    }
}

fn header_value_bytes(value: &str) -> Value {
    let bytes = value
        .as_bytes()
        .iter()
        .map(|byte| Value::F64(*byte as f64))
        .collect::<Vec<_>>();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(bytes)))
}

fn header_values_array(values: &[Vec<u8>]) -> Value {
    let arrays = values
        .iter()
        .map(|value| header_value_bytes(&String::from_utf8_lossy(value)))
        .collect::<Vec<_>>();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(arrays)))
}

fn header_entries_array(entries: &[(String, Vec<u8>)]) -> Value {
    let pairs = entries
        .iter()
        .map(|(name, value)| {
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                Value::String(Arc::from(name.as_str())),
                header_value_bytes(&String::from_utf8_lossy(value)),
            ])))
        })
        .collect::<Vec<_>>();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(pairs)))
}

#[derive(Debug, Clone)]
struct HttpResponseData {
    status: u16,
    headers: Vec<(String, Vec<u8>)>,
    body: String,
}

fn map_transport_error(message: &str) -> &'static str {
    let lowercase = message.to_lowercase();
    if lowercase.contains("refused") {
        "connection-refused"
    } else if lowercase.contains("timed out") {
        "connection-timeout"
    } else if lowercase.contains("invalid") {
        "HTTP-request-URI-invalid"
    } else {
        "internal-error"
    }
}

fn http_request(method: &str, url: &str, body: Option<&str>) -> Result<HttpResponseData, String> {
    let url = if url.starts_with("http://") {
        &url[7..]
    } else if url.starts_with("https://") {
        return Err("HTTPS not supported (use http://)".into());
    } else {
        url
    };

    let (host_port, path) = match url.find('/') {
        Some(index) => (&url[..index], &url[index..]),
        None => (url, "/"),
    };

    let (host, port) = match host_port.find(':') {
        Some(index) => (
            &host_port[..index],
            host_port[index + 1..].parse::<u16>().unwrap_or(80),
        ),
        None => (host_port, 80u16),
    };

    let mut stream = std::net::TcpStream::connect((host, port))
        .map_err(|error| format!("Connection failed: {}", error))?;

    stream
        .set_read_timeout(Some(std::time::Duration::from_secs(10)))
        .map_err(|error| format!("Timeout config failed: {}", error))?;

    let content_length = body.map(|text| text.len()).unwrap_or(0);
    let mut request = format!(
        "{} {} HTTP/1.1\r\nHost: {}\r\nConnection: close\r\nContent-Length: {}\r\n\r\n",
        method, path, host, content_length,
    );
    if let Some(body) = body {
        request.push_str(body);
    }

    stream
        .write_all(request.as_bytes())
        .map_err(|error| format!("Write failed: {}", error))?;

    let mut raw_response = String::new();
    stream
        .read_to_string(&mut raw_response)
        .map_err(|error| format!("Read failed: {}", error))?;

    let (head, body) = match raw_response.find("\r\n\r\n") {
        Some(index) => (
            &raw_response[..index],
            raw_response[index + 4..].to_string(),
        ),
        None => (raw_response.as_str(), String::new()),
    };

    let mut lines = head.lines();
    let status = lines
        .next()
        .and_then(|status_line| status_line.split_whitespace().nth(1))
        .and_then(|code| code.parse::<u16>().ok())
        .unwrap_or(0);

    let headers = lines
        .filter_map(|line| line.split_once(':'))
        .map(|(name, value)| (name.trim().to_string(), value.trim().as_bytes().to_vec()))
        .collect::<Vec<_>>();

    Ok(HttpResponseData {
        status,
        headers,
        body,
    })
}

fn register_resource_types(vm: &mut VM) -> HttpTypeIds {
    fn resource(vm: &mut VM, export_name: &str, type_name: &str) -> usize {
        let mut type_def = TypeDef::new(type_name);
        type_def.interface = Some("wasi:http/types".into());
        type_def.is_resource = true;
        let type_id = vm.type_registry.register(type_def);
        vm.register_host_resource_type_export("wasi:http/types", export_name, type_id);
        type_id
    }

    HttpTypeIds {
        headers: resource(vm, "headers", "HttpHeaders"),
        outgoing_request: resource(vm, "outgoing-request", "HttpOutgoingRequest"),
        request_options: resource(vm, "request-options", "HttpRequestOptions"),
        incoming_response: resource(vm, "incoming-response", "HttpIncomingResponse"),
        future_incoming_response: resource(
            vm,
            "future-incoming-response",
            "HttpFutureIncomingResponse",
        ),
        incoming_body: resource(vm, "incoming-body", "HttpIncomingBody"),
        future_trailers: resource(vm, "future-trailers", "HttpFutureTrailers"),
        outgoing_response: resource(vm, "outgoing-response", "HttpOutgoingResponse"),
        outgoing_body: resource(vm, "outgoing-body", "HttpOutgoingBody"),
        _incoming_request: resource(vm, "incoming-request", "HttpIncomingRequest"),
        _response_outparam: resource(vm, "response-outparam", "HttpResponseOutparam"),
    }
}

/// `wasi:http/types` — the four resources `wasi:http@0.3.1` declares.
///
/// THIRTY-NINE registrations were removed from here on 2026-08-21. 0.2 had
/// thirteen resources where 0.3.1 has four, and every one of the extras was
/// still bound:
///
///   * `outgoing-request` / `incoming-request` → one `request`
///   * `outgoing-response` / `incoming-response` → one `response`
///   * `future-incoming-response` → gone; `client.send` answers the response
///   * `outgoing-body` / `incoming-body` / `future-trailers` → gone; a body is
///     a `stream<u8>` passed to `new` and taken back by `consume-body`
///   * `response-outparam` → gone; `handler.handle` RETURNS its response
///   * `http-error-code` → gone with the `wasi:io` error it borrowed
///   * `request-options.{connect,first-byte,between-bytes}-timeout` → the
///     getters gained a `get-` prefix
///
/// The registry still keeps requests and responses in two tables each, because
/// a request built here and a request off the wire genuinely differ in what
/// they can carry. 0.3.1 constrains the INTERFACE, not the storage: the
/// accessors read both tables (see `register_wasi3_accessors`), which is what
/// makes them one resource from the guest's side.
///
/// Nothing was kept "for compatibility". A 0.2 name that still resolves fails
/// nowhere — the caller keeps working against this host and breaks only against
/// a conforming one, with nothing left pointing at why. `interface_coverage.rs`
/// asserts registered ⊆ spec so a re-introduction fails by name.
fn register_types(vm: &mut VM, type_ids: HttpTypeIds) {
    vm.register_host_fn(
        "wasi:http/types",
        "[constructor]fields",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            let mut registry = registry().lock().unwrap();
            let id = alloc_id();
            registry.headers.insert(
                id,
                HeadersResource {
                    entries: Vec::new(),
                },
            );
            make_resource(KIND_HEADERS, id, type_ids.headers)
        }),
    );

    // `[method]fields.entries` USED TO BE REGISTERED HERE.
    //
    // 0.3.1 calls it `copy-all`, registered below with the same body. `entries`
    // is not declared by `wasi:http@0.3.1` at all, so a guest built against the
    // WIT cannot ask for it and a conforming runtime cannot answer it.

    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.has",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let Some(name) = string_arg(args, 1) else {
                return Value::Bool(false);
            };
            let registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get(&headers_id) else {
                return err("invalid-argument");
            };
            let target = name.to_ascii_lowercase();
            Value::Bool(
                headers
                    .entries
                    .iter()
                    .any(|(entry_name, _)| entry_name.eq_ignore_ascii_case(&target)),
            )
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.get",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let Some(name) = string_arg(args, 1) else {
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())));
            };
            let registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get(&headers_id) else {
                return err("invalid-argument");
            };
            let target = name.to_ascii_lowercase();
            let values = headers
                .entries
                .iter()
                .filter_map(|(entry_name, value)| {
                    if entry_name.eq_ignore_ascii_case(&target) {
                        Some(value.clone())
                    } else {
                        None
                    }
                })
                .collect::<Vec<_>>();
            header_values_array(&values)
        }),
    );







    vm.register_host_fn(
        "wasi:http/types",
        "[constructor]request-options",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            let mut registry = registry().lock().unwrap();
            let id = alloc_id();
            registry.request_options.insert(
                id,
                RequestOptionsResource {
                    connect_timeout_ns: None,
                    first_byte_timeout_ns: None,
                    between_bytes_timeout_ns: None,
                },
            );
            make_resource(KIND_REQUEST_OPTIONS, id, type_ids.request_options)
        }),
    );




    // future-incoming-response.subscribe → pollable (always ready in sync model)

    // fields.set(name, values) → result
    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.set",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let Some(name) = string_arg(args, 1) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get_mut(&headers_id) else {
                return err("invalid-argument");
            };
            let key = name.to_ascii_lowercase();
            // §fields.set replaces every value the field had.
            let previous = headers.entries.len();
            headers
                .entries
                .retain(|(k, _)| !k.eq_ignore_ascii_case(&key));
            // `value` is `list<field-value>` in the WIT, but a single value is
            // how field values travel everywhere else in this module —
            // `fields.append` takes one, and `push_incoming_request` carries
            // `(String, Vec<u8>)`. Accepting ONLY an array is what this did
            // until 2026-08-21: `fields.set(h, "x", "yes")` removed the old
            // values, added nothing, and answered `Value::Null` — success. The
            // header vanished and no caller could tell. Every test of `set`
            // asserted the call did not error, which it never did.
            match args.get(2) {
                Some(Value::Object(object)) => {
                    let inner = object.lock().unwrap();
                    match inner.kind {
                        vybe_runtime::value::ObjectKind::Array(ref elems) => {
                            for value in elems {
                                headers
                                    .entries
                                    .push((key.clone(), format!("{}", value).into_bytes()));
                            }
                        }
                        // Not a list and not a scalar: there is no reading of
                        // this that sets a header, so it is an error rather
                        // than a removal the caller did not ask for.
                        _ => {
                            drop(inner);
                            drop(registry);
                            return err("invalid-syntax");
                        }
                    }
                }
                // §set with an EMPTY list removes the field, so an absent
                // argument is the one case where writing nothing is correct.
                None | Some(Value::Null) => {
                    let _ = previous;
                }
                Some(single) => headers
                    .entries
                    .push((key.clone(), format!("{}", single).into_bytes())),
            }
            Value::Null
        }),
    );

    // fields.delete(name) → result
    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.delete",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let Some(name) = string_arg(args, 1) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get_mut(&headers_id) else {
                return err("invalid-argument");
            };
            headers
                .entries
                .retain(|(k, _)| !k.eq_ignore_ascii_case(&name));
            Value::Null
        }),
    );

    // fields.append(name, value) → result
    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.append",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let Some(name) = string_arg(args, 1) else {
                return err("invalid-argument");
            };
            let value = string_arg(args, 2).unwrap_or_default().into_bytes();
            let mut registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get_mut(&headers_id) else {
                return err("invalid-argument");
            };
            headers.entries.push((name.to_ascii_lowercase(), value));
            Value::Null
        }),
    );

    // [method]fields.clone → new fields resource with same entries
    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.clone",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(src_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let entries = registry
                .headers
                .get(&src_id)
                .map(|h| h.entries.clone())
                .unwrap_or_default();
            let new_id = alloc_id();
            registry.headers.insert(new_id, HeadersResource { entries });
            make_resource(KIND_HEADERS, new_id, type_ids.headers)
        }),
    );

    // [static]fields.from-list(entries) → result<fields, header-error>
    vm.register_host_fn(
        "wasi:http/types",
        "[static]fields.from-list",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
            if let Some(Value::Object(arr)) = args.first() {
                let inner = arr.lock().unwrap();
                if let vybe_runtime::value::ObjectKind::Array(ref pairs) = inner.kind {
                    for pair in pairs {
                        if let Value::Object(p) = pair {
                            let p = p.lock().unwrap();
                            if let vybe_runtime::value::ObjectKind::Array(ref kv) = p.kind {
                                let name = kv
                                    .first()
                                    .map(|v| format!("{}", v))
                                    .unwrap_or_default()
                                    .to_ascii_lowercase();
                                let val = kv
                                    .get(1)
                                    .map(|v| format!("{}", v).into_bytes())
                                    .unwrap_or_default();
                                entries.push((name, val));
                            }
                        }
                    }
                }
            }
            let mut registry = registry().lock().unwrap();
            let id = alloc_id();
            registry.headers.insert(id, HeadersResource { entries });
            make_resource(KIND_HEADERS, id, type_ids.headers)
        }),
    );

    // outgoing-request getters




    // [method]outgoing-request.body → result<outgoing-body, error-code>

    // request-options timeout getters/setters (durations in nanoseconds)

    vm.register_host_fn(
        "wasi:http/types",
        "[method]request-options.set-connect-timeout",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let ns = args.get(1).map(|v| v.as_f64() as u64);
            let mut registry = registry().lock().unwrap();
            let Some(opt) = registry.request_options.get_mut(&id) else {
                return err("invalid-argument");
            };
            opt.connect_timeout_ns = ns;
            Value::Null
        }),
    );


    vm.register_host_fn(
        "wasi:http/types",
        "[method]request-options.set-first-byte-timeout",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let ns = args.get(1).map(|v| v.as_f64() as u64);
            let mut registry = registry().lock().unwrap();
            let Some(opt) = registry.request_options.get_mut(&id) else {
                return err("invalid-argument");
            };
            opt.first_byte_timeout_ns = ns;
            Value::Null
        }),
    );


    vm.register_host_fn(
        "wasi:http/types",
        "[method]request-options.set-between-bytes-timeout",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let ns = args.get(1).map(|v| v.as_f64() as u64);
            let mut registry = registry().lock().unwrap();
            let Some(opt) = registry.request_options.get_mut(&id) else {
                return err("invalid-argument");
            };
            opt.between_bytes_timeout_ns = ns;
            Value::Null
        }),
    );

    // incoming-response.consume → incoming-body

    // incoming-body.%stream → input-stream resource readable via [method]input-stream.blocking-read
    // Registers the body bytes as a Buffer stream in the filesystem registry so standard
    // stream host functions can drain it without any bespoke logic in http.rs.

    // [static]incoming-body.finish → future-trailers

    // future-trailers.get → option<result<option<trailers>, error-code>>

    // future-trailers.subscribe → always-ready pollable

    // [constructor]outgoing-response(fields) → outgoing-response




    // [method]outgoing-response.body → result<outgoing-body, error-code>

    // [method]outgoing-body.write → result<output-stream, error-code>
    // Returns a simplified stream object with __body_id for later flush.

    // [static]outgoing-body.finish → result<_, error-code>

    // http-error-code(err) → option<error-code>

    // ── incoming-request (server side) ─────────────────────────────────────
    //
    // wasi:http/types §incoming-request. The host builds one of these from the
    // transport it is serving via `push_incoming_request`; guest code reads it
    // through these accessors. Previously all six returned `Value::Null` as
    // link padding.


    // `path-with-query`, `scheme` and `authority` are `option<...>` in the WIT:
    // absent is null, not an error.




    // §incoming-request.consume: succeeds at most ONCE; later calls are errors.

    // ── response-outparam (server side) ────────────────────────────────────
    //
    // §response-outparam.set: "Set the value of the `response-outparam` to
    // either send a response, or indicate an error. This method consumes the
    // `response-outparam` to ensure that it is called at most once."

    // §response-outparam.send-informational — 1xx interim responses.
}

// ── Host-side entry points ──────────────────────────────────────────────────
//
// The transport (hyper, in `vybex --serve`) has no WIT constructor for
// `incoming-request` or `response-outparam` — per the spec those are produced
// BY the host and handed TO the guest. These are that seam.

/// Create an `incoming-request` resource and return its id.
pub fn push_incoming_request(
    method: &str,
    path_with_query: Option<String>,
    scheme: Option<String>,
    authority: Option<String>,
    headers: Vec<(String, Vec<u8>)>,
    body: Vec<u8>,
) -> u32 {
    let mut registry = registry().lock().unwrap();
    let headers_id = alloc_id();
    registry
        .headers
        .insert(headers_id, HeadersResource { entries: headers });
    let id = alloc_id();
    registry.incoming_requests.insert(
        id,
        IncomingRequestResource {
            method: method.to_string(),
            path_with_query,
            scheme,
            authority,
            headers_id,
            body,
            consumed: false,
        },
    );
    id
}

/// Create a `response-outparam` for a request and return its id.
pub fn push_response_outparam() -> u32 {
    let mut registry = registry().lock().unwrap();
    let id = alloc_id();
    registry
        .response_outparams
        .insert(id, ResponseOutparamResource::default());
    id
}

/// What the guest set on a `response-outparam`: `Ok((status, headers, body))`
/// or `Err(error-code)`. `None` if the guest never called `set`.
pub type ResponseParts = (u16, Vec<(String, Vec<u8>)>, Vec<u8>);

pub fn take_response_outparam(id: u32) -> Option<Result<ResponseParts, String>> {
    let mut registry = registry().lock().unwrap();
    let param = registry.response_outparams.remove(&id)?;
    if !param.set {
        return None;
    }
    if let Some(error) = param.error {
        return Some(Err(error));
    }
    let response_id = param.response_id?;
    let response = registry.outgoing_responses.get(&response_id)?.clone();
    let headers = registry
        .headers
        .get(&response.headers_id)
        .map(|h| h.entries.clone())
        .unwrap_or_default();
    let body = response
        .body_id
        .and_then(|bid| registry.outgoing_bodies.get(&bid))
        .map(|b| b.bytes.clone())
        .unwrap_or_default();
    Some(Ok((response.status, headers, body)))
}

/// Create the `response` a served request will answer with, host-side.
///
/// 0.2 handed the guest a `response-outparam` to write INTO. 0.3.1 has no such
/// resource: `handler.handle` returns its response, so the response itself is
/// the only thing there is to hand over. The host builds it empty — status
/// 200, no headers, per §response.new's default — and the guest mutates it
/// through `response.set-status-code` and `response.get-headers`.
///
/// Pre-creating it rather than waiting for the guest to call `response.new` is
/// what keeps a served request to ONE response: a script that sets nothing
/// still answers through this resource rather than through a second path.
pub fn push_response() -> u32 {
    let mut registry = registry().lock().unwrap();
    let headers_id = alloc_id();
    registry
        .headers
        .insert(headers_id, HeadersResource { entries: Vec::new() });
    let id = alloc_id();
    registry.outgoing_responses.insert(
        id,
        OutgoingResponseResource {
            status: 200,
            headers_id,
            body_id: None,
        },
    );
    id
}

/// The guest-facing handle for [`push_response`].
pub fn response_value(vm: &VM, response_id: u32) -> Option<Value> {
    let type_id = vm.type_registry.get_id("HttpOutgoingResponse")?;
    Some(make_resource(KIND_OUTGOING_RESPONSE, response_id, type_id))
}

/// Read a `response` resource back into its parts, host-side.
///
/// This is what replaces `response-outparam` for a HOST that needs the guest's
/// answer. 0.3.1 deleted the outparam because `handler.handle` RETURNS its
/// response; a host with no export to call — `vybex --serve` compiles a script,
/// not a component — still needs some way to take delivery, and the resource
/// itself is that way. Reading the response is not a `wasi:` function and is
/// not registered as one: it is the host side of the boundary, the same role
/// `push_incoming_request` plays for the request.
///
/// Accepts a response from either direction, because 0.3.1 has one `response`.
pub fn take_response(value: &Value) -> Option<ResponseParts> {
    let registry = registry().lock().unwrap();
    let (status, headers_id, body) = if let Some(id) =
        resource_id(value, KIND_OUTGOING_RESPONSE)
    {
        let response = registry.outgoing_responses.get(&id)?;
        let body = response
            .body_id
            .and_then(|body_id| registry.outgoing_bodies.get(&body_id))
            .map(|body| body.bytes.clone())
            // A response whose body was never opened has an EMPTY body, which
            // is a legal answer — not a missing resource.
            .unwrap_or_default();
        (response.status, response.headers_id, body)
    } else {
        let id = resource_id(value, KIND_INCOMING_RESPONSE)?;
        let response = registry.incoming_responses.get(&id)?;
        (
            response.status,
            response.headers_id,
            response.body.as_bytes().to_vec(),
        )
    };
    let headers = registry
        .headers
        .get(&headers_id)
        .map(|headers| headers.entries.clone())
        .unwrap_or_default();
    Some((status, headers, body))
}

/// Build the guest-facing resource values for a served request.
pub fn incoming_request_value(vm: &VM, request_id: u32) -> Option<Value> {
    let type_id = vm.type_registry.get_id("HttpIncomingRequest")?;
    Some(make_resource(KIND_INCOMING_REQUEST, request_id, type_id))
}

pub fn response_outparam_value(vm: &VM, param_id: u32) -> Option<Value> {
    let type_id = vm.type_registry.get_id("HttpResponseOutparam")?;
    Some(make_resource(KIND_RESPONSE_OUTPARAM, param_id, type_id))
}

// `wasi:http/outgoing-handler.handle` USED TO BE REGISTERED HERE.
//
// The interface does not exist in `wasi:http@0.3.1` — `worlds.wit` and
// `types.wit` declare `client` and `handler`, and nothing else. 0.3 replaced
// 0.2's two-step (`handle` -> `future-incoming-response` -> `.get`) with a
// single `client.send` that answers the response directly, and
// `register_wasi3_handler` below already implemented it.
//
// Keeping both bound meant every caller in the tree could stay on the 0.2
// shape indefinitely without anything failing, which is what happened: the
// four transport test files and one interop corpus file were still written
// against the future when this was removed on 2026-08-21.

pub fn register(vm: &mut VM) {
    let type_ids = register_resource_types(vm);
    register_types(vm, type_ids);
    register_wasi3(vm, type_ids);
    register_wasi3_handler(vm, type_ids);
    register_wasi3_accessors(vm, type_ids);
}

/// WASI 0.3 accessors for the `request` / `response` resources.
///
/// 0.3 renamed both the resources and their getters relative to 0.2:
///   * `outgoing-request` → `request`, `incoming-response` → `response`
///   * bare getters gained a `get-` prefix (`method` → `get-method`,
///     `status` → `get-status-code`, `headers` → `get-headers`, …)
///
/// The underlying resources are shared with the 0.2 surface (a request made by
/// `[static]request.new` is the same `KIND_OUTGOING_REQUEST` a 0.2 constructor
/// produces), so these are spec-named views over the same registry state, not
/// a second implementation.
fn register_wasi3_accessors(vm: &mut VM, type_ids: HttpTypeIds) {
    // ── request getters ────────────────────────────────────────────────
    //
    // 0.3.1 declares ONE `request` resource where 0.2 had `outgoing-request`
    // and `incoming-request`, so a `request.get-*` accessor has to answer for
    // a request arriving from the network as readily as for one built by
    // `request.new`. This host still keeps the two directions in separate
    // tables; the view below is what makes them one resource from the guest's
    // side, the same way `response.get-status-code` already reads both
    // response tables.
    //
    // Reading through a view rather than widening `resource_id` keeps the
    // asymmetry that IS real: an incoming request has no `request-options`,
    // because nothing on this side chose them.
    struct RequestView {
        method: String,
        path_with_query: Option<String>,
        scheme: Option<String>,
        authority: Option<String>,
        headers_id: u32,
        options_id: Option<u32>,
    }

    fn with_request<T>(args: &[Value], f: impl FnOnce(&RequestView) -> T) -> Option<T> {
        let registry = registry().lock().unwrap();
        if let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) {
            let request = registry.outgoing_requests.get(&id)?;
            return Some(f(&RequestView {
                method: request.method.clone(),
                path_with_query: request.path_with_query.clone(),
                scheme: request.scheme.clone(),
                authority: request.authority.clone(),
                headers_id: request.headers_id,
                options_id: request.options_id,
            }));
        }
        let id = resource_id(&args[0], KIND_INCOMING_REQUEST)?;
        let request = registry.incoming_requests.get(&id)?;
        Some(f(&RequestView {
            method: request.method.clone(),
            path_with_query: request.path_with_query.clone(),
            scheme: request.scheme.clone(),
            authority: request.authority.clone(),
            headers_id: request.headers_id,
            options_id: None,
        }))
    }

    vm.register_host_fn(
        "wasi:http/types",
        "[method]request.get-method",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match with_request(args, |r| r.method.clone()) {
                Some(method) => Value::String(Arc::from(method.as_str())),
                None => err("invalid-argument"),
            }
        }),
    );
    for (name, pick) in [
        (
            "[method]request.get-path-with-query",
            (|r: &RequestView| r.path_with_query.clone()) as fn(&_) -> Option<String>,
        ),
        ("[method]request.get-scheme", |r| r.scheme.clone()),
        ("[method]request.get-authority", |r| r.authority.clone()),
    ] {
        vm.register_host_fn(
            "wasi:http/types",
            name,
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                // `option<string>` — absent is null.
                match with_request(args, pick) {
                    Some(Some(value)) => Value::String(Arc::from(value.as_str())),
                    Some(None) => Value::Null,
                    None => err("invalid-argument"),
                }
            }),
        );
    }
    vm.register_host_fn(
        "wasi:http/types",
        "[method]request.get-headers",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match with_request(args, |r| r.headers_id) {
                Some(headers_id) => make_resource(KIND_HEADERS, headers_id, type_ids.headers),
                None => err("invalid-argument"),
            }
        }),
    );
    // `get-options` — request-options are not retained per request (the 0.3
    // `request.new` accepts them but the transport applies no timeouts), so
    // the option is always absent rather than fabricated.
    vm.register_host_fn(
        "wasi:http/types",
        "[method]request.get-options",
        // §request.get-options -> option<request-options>: the options handed
        // to `request.new`, or none.
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match with_request(args, |r| r.options_id) {
                Some(Some(options_id)) => {
                    make_resource(KIND_REQUEST_OPTIONS, options_id, type_ids.request_options)
                }
                Some(None) => Value::Null,
                None => err("invalid-argument"),
            }
        }),
    );

    // ── fields (0.3 additions) ─────────────────────────────────────────
    // `copy-all` is 0.3's name for the full name/value list (0.2: `entries`).
    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.copy-all",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            match registry.headers.get(&headers_id) {
                Some(headers) => header_entries_array(&headers.entries),
                None => err("invalid-argument"),
            }
        }),
    );
    // `get-and-delete(name) -> list<field-value>` — read every value for the
    // name, then remove them all (case-insensitive, per field-name matching).
    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.get-and-delete",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let Some(name) = string_arg(args, 1) else {
                return Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())));
            };
            let target = name.to_ascii_lowercase();
            let mut registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get_mut(&headers_id) else {
                return err("invalid-argument");
            };
            let values = headers
                .entries
                .iter()
                .filter(|(key, _)| key.to_ascii_lowercase() == target)
                .map(|(_, value)| value.clone())
                .collect::<Vec<_>>();
            headers
                .entries
                .retain(|(key, _)| key.to_ascii_lowercase() != target);
            header_values_array(&values)
        }),
    );

    // ── request-options (0.3 `get-` prefixed names + clone) ────────────
    for (name, pick) in [
        (
            "[method]request-options.get-connect-timeout",
            (|o: &RequestOptionsResource| o.connect_timeout_ns) as fn(&_) -> Option<u64>,
        ),
        ("[method]request-options.get-first-byte-timeout", |o| {
            o.first_byte_timeout_ns
        }),
        ("[method]request-options.get-between-bytes-timeout", |o| {
            o.between_bytes_timeout_ns
        }),
    ] {
        vm.register_host_fn(
            "wasi:http/types",
            name,
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                    return err("invalid-argument");
                };
                let registry = registry().lock().unwrap();
                match registry.request_options.get(&id) {
                    // `option<duration>` — absent is null.
                    Some(options) => pick(options)
                        .map(|ns| Value::F64(ns as f64))
                        .unwrap_or(Value::Null),
                    None => err("invalid-argument"),
                }
            }),
        );
    }
    vm.register_host_fn(
        "wasi:http/types",
        "[method]request-options.clone",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(source) = registry.request_options.get(&id).cloned() else {
                return err("invalid-argument");
            };
            let new_id = alloc_id();
            registry.request_options.insert(new_id, source);
            drop(registry);
            make_resource(KIND_REQUEST_OPTIONS, new_id, type_ids.request_options)
        }),
    );

    // ── request setters (0.3 names over the shared resource) ───────────
    vm.register_host_fn(
        "wasi:http/types",
        "[method]request.set-method",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let Some(method) = string_arg(args, 1) else {
                return err("HTTP-request-method-invalid");
            };
            if method.trim().is_empty() {
                return err("HTTP-request-method-invalid");
            }
            let mut registry = registry().lock().unwrap();
            match registry.outgoing_requests.get_mut(&id) {
                Some(request) => {
                    request.method = method.trim().to_ascii_uppercase();
                    Value::Null
                }
                None => err("invalid-argument"),
            }
        }),
    );
    // `option<string>` setters — a null argument clears the field.
    for (name, apply) in [
        (
            "[method]request.set-path-with-query",
            (|r: &mut OutgoingRequestResource, v: Option<String>| r.path_with_query = v)
                as fn(&mut _, Option<String>),
        ),
        // RFC 9110 §4.2: a scheme is case-insensitive and is normalised to
        // lowercase. `client.send` compares it with `!=`, so a request whose
        // scheme was stored as given is rejected for a scheme that is
        // perfectly valid. The 0.2 setter this replaced normalised here; the
        // rename dropped it, which is a behaviour lost to a name change and
        // exactly what the corpus test for `HtTp` exists to catch.
        ("[method]request.set-scheme", |r, v| {
            r.scheme = v.map(|scheme| scheme.to_ascii_lowercase())
        }),
        ("[method]request.set-authority", |r, v| r.authority = v),
    ] {
        vm.register_host_fn(
            "wasi:http/types",
            name,
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                    return err("invalid-argument");
                };
                let value = string_arg(args, 1);
                let mut registry = registry().lock().unwrap();
                match registry.outgoing_requests.get_mut(&id) {
                    Some(request) => {
                        apply(request, value);
                        Value::Null
                    }
                    None => err("invalid-argument"),
                }
            }),
        );
    }

    // ── response getters / setters ─────────────────────────────────────
    vm.register_host_fn(
        "wasi:http/types",
        "[method]response.get-status-code",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) {
                let registry = registry().lock().unwrap();
                return match registry.incoming_responses.get(&id) {
                    Some(response) => Value::I32(response.status as i32),
                    None => err("invalid-argument"),
                };
            }
            if let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) {
                let registry = registry().lock().unwrap();
                return match registry.outgoing_responses.get(&id) {
                    Some(response) => Value::I32(response.status as i32),
                    None => err("invalid-argument"),
                };
            }
            err("invalid-argument")
        }),
    );
    vm.register_host_fn(
        "wasi:http/types",
        "[method]response.set-status-code",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) else {
                return err("invalid-argument");
            };
            let status = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(0);
            if !(100..=599).contains(&status) {
                return err("invalid-argument");
            }
            let mut registry = registry().lock().unwrap();
            match registry.outgoing_responses.get_mut(&id) {
                Some(response) => {
                    response.status = status as u16;
                    Value::Null
                }
                None => err("invalid-argument"),
            }
        }),
    );
    vm.register_host_fn(
        "wasi:http/types",
        "[method]response.get-headers",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let headers_id = if let Some(id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) {
                registry()
                    .lock()
                    .unwrap()
                    .incoming_responses
                    .get(&id)
                    .map(|r| r.headers_id)
            } else if let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) {
                registry()
                    .lock()
                    .unwrap()
                    .outgoing_responses
                    .get(&id)
                    .map(|r| r.headers_id)
            } else {
                None
            };
            match headers_id {
                Some(headers_id) => make_resource(KIND_HEADERS, headers_id, type_ids.headers),
                None => err("invalid-argument"),
            }
        }),
    );
}

/// `wasi:http@0.3.1` `client` + `handler` interfaces.
///
/// 0.3 collapses the 0.2 dance (`outgoing-handler.handle` → `future-incoming-
/// response` → `.get()`) into a single async call that yields the response
/// directly:
///
/// ```wit
/// interface client  { send:   async func(request) -> result<response, error-code>; }
/// interface handler { handle: async func(request) -> result<response, error-code>; }
/// ```
///
/// `client.send` and `handler.handle` are intentionally identical in signature
/// (per the spec note: WIT can't represent importing two instances of the same
/// interface, so `client` duplicates `handler`). Both are wired to the same
/// implementation here.
///
/// Async is executed synchronously host-side: the result is the resolved
/// `incoming-response` resource, or an `error-code` on transport failure.
fn register_wasi3_handler(vm: &mut VM, type_ids: HttpTypeIds) {
    for (module, name) in [
        ("wasi:http/client", "send"),
        ("wasi:http/handler", "handle"),
    ] {
        vm.register_host_fn(
            module,
            name,
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                    return err("HTTP-request-denied");
                };

                let request = {
                    let registry = registry().lock().unwrap();
                    let Some(request) = registry.outgoing_requests.get(&request_id) else {
                        return err("HTTP-request-denied");
                    };
                    request.clone()
                };

                let scheme = request.scheme.unwrap_or_else(|| "http".into());
                if scheme != "http" {
                    return err("HTTP-request-URI-invalid");
                }
                let Some(authority) = request
                    .authority
                    .filter(|authority| !authority.trim().is_empty())
                else {
                    return err("HTTP-request-URI-invalid");
                };

                let mut path = request.path_with_query.unwrap_or_else(|| "/".into());
                if !path.starts_with('/') {
                    path = format!("/{}", path);
                }
                let url = format!("{}://{}{}", scheme, authority, path);

                match http_request(&request.method, &url, None) {
                    Ok(response) => {
                        let mut registry = registry().lock().unwrap();
                        let headers_id = alloc_id();
                        registry.headers.insert(
                            headers_id,
                            HeadersResource {
                                entries: response.headers,
                            },
                        );
                        let response_id = alloc_id();
                        registry.incoming_responses.insert(
                            response_id,
                            IncomingResponseResource {
                                status: response.status,
                                headers_id,
                                body: response.body,
                            },
                        );
                        drop(registry);
                        make_resource(
                            KIND_INCOMING_RESPONSE,
                            response_id,
                            type_ids.incoming_response,
                        )
                    }
                    Err(message) => err(map_transport_error(&message)),
                }
            }),
        );
    }
}

fn register_wasi3(vm: &mut VM, type_ids: HttpTypeIds) {
    // [static]request.new — WASI 0.3.1 constructor.
    vm.register_host_fn(
        "wasi:http/types",
        "[static]request.new",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            // `headers` is an OWNED `fields`. Defaulting a bad handle to id 0
            // (which this did until 2026-08-21) turns a caller's mistake into
            // "a request with headers it never set" — the 0.2 constructor this
            // replaced answered `invalid-argument`, and losing that to a rename
            // is how a wrong call starts succeeding.
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            // §request.new(headers, contents, trailers, options) — the 4th
            // argument is `option<request-options>`, read back by
            // `request.get-options`. Accept it in either the spec position or
            // immediately after headers, since callers that pass no body still
            // want options to land.
            let options_id = args
                .get(3)
                .and_then(|a| resource_id(a, KIND_REQUEST_OPTIONS))
                .or_else(|| {
                    args.get(1)
                        .and_then(|a| resource_id(a, KIND_REQUEST_OPTIONS))
                });
            // `contents` sits in the spec's second position. When a caller
            // put request-options there instead, that is not a stream and
            // drains to nothing — the same answer an absent body gives.
            let body = if options_id.is_some()
                && args
                    .get(1)
                    .and_then(|a| resource_id(a, KIND_REQUEST_OPTIONS))
                    .is_some()
            {
                Vec::new()
            } else {
                args.get(1).map(|c| ctx.stream_drain(c)).unwrap_or_default()
            };
            let mut reg = registry().lock().unwrap();
            let id = alloc_id();
            reg.outgoing_requests.insert(
                id,
                OutgoingRequestResource {
                    options_id,
                    headers_id,
                    body,
                    method: "GET".into(),
                    path_with_query: None,
                    scheme: None,
                    authority: None,
                },
            );
            make_resource(KIND_OUTGOING_REQUEST, id, type_ids.outgoing_request)
        }),
    );

    // [static]request.consume-body — returns body bytes as stream<u8> + resolved trailers future.
    vm.register_host_fn(
        "wasi:http/types",
        "[static]request.consume-body",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match consume_body_bytes(&args[0]) {
            Some(bytes) => body_stream_tuple(ctx, bytes),
            None => err("invalid-argument"),
        }),
    );

    // [static]response.new — WASI 0.3 constructor alongside [constructor]outgoing-response.
    vm.register_host_fn(
        "wasi:http/types",
        "[static]response.new",
        // §response.new(headers, contents, trailers). `contents` is a
        // `stream<u8>` and IS the body — 0.3.1 has no `outgoing-body` resource
        // and no `body`/`write`/`finish` calls, so this is the only way a
        // response body can come into existence. Dropping the argument (which
        // this did until 2026-08-21, behind a comment claiming host draining
        // was unavailable) meant every response was silently empty.
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            // Drained here rather than lazily: `stream_drain` needs the VM, and
            // the guest is free to drop its end the moment `new` returns.
            // A non-stream argument drains to nothing, which is what an absent
            // `option<stream>` should read as anyway.
            let body = args.get(1).map(|c| ctx.stream_drain(c)).unwrap_or_default();
            let trailers = args
                .get(2)
                .and_then(|t| resource_id(t, KIND_HEADERS))
                .and_then(|id| {
                    registry()
                        .lock()
                        .unwrap()
                        .headers
                        .get(&id)
                        .map(|h| h.entries.clone())
                })
                .unwrap_or_default();
            let mut reg = registry().lock().unwrap();
            let body_id = if body.is_empty() && trailers.is_empty() {
                // No body is not the same as an EMPTY body only if something
                // downstream can tell them apart; nothing here can, so the
                // cheaper representation is the honest one.
                None
            } else {
                let body_id = alloc_id();
                reg.outgoing_bodies.insert(
                    body_id,
                    OutgoingBodyResource {
                        bytes: body,
                        finished: true,
                        trailers,
                    },
                );
                Some(body_id)
            };
            let id = alloc_id();
            reg.outgoing_responses.insert(
                id,
                OutgoingResponseResource {
                    status: 200,
                    headers_id,
                    body_id,
                },
            );
            make_resource(KIND_OUTGOING_RESPONSE, id, type_ids.outgoing_response)
        }),
    );

    // [static]response.consume-body — returns body bytes as stream<u8> + resolved trailers future.
    vm.register_host_fn(
        "wasi:http/types",
        "[static]response.consume-body",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match consume_body_bytes(&args[0]) {
            Some(bytes) => body_stream_tuple(ctx, bytes),
            None => err("invalid-argument"),
        }),
    );
}

/// The bytes a 0.3 `consume-body` should hand over, for any resource its
/// signature can legally receive. `None` means "not a resource this call
/// accepts, or already consumed" — an `error-code`, not an empty body.
///
/// §request.consume-body and §response.consume-body take the REQUEST or the
/// RESPONSE, not its body: 0.3 deleted the `incoming-body`/`outgoing-body`
/// resources that 0.2 made you fetch first. Both registrations used to match
/// only the two BODY kinds, so a caller passing the resource the signature
/// actually names — which is every conforming caller — fell through to
/// `Vec::new()` and read an empty body with no error raised anywhere. The body
/// kinds stay accepted for callers still holding a 0.2-shaped handle.
fn consume_body_bytes(arg: &Value) -> Option<Vec<u8>> {
    if let Some(body_id) = resource_id(arg, KIND_OUTGOING_BODY) {
        let registry = registry().lock().unwrap();
        return registry.outgoing_bodies.get(&body_id).map(|b| b.bytes.clone());
    }
    if let Some(body_id) = resource_id(arg, KIND_INCOMING_BODY) {
        let registry = registry().lock().unwrap();
        return registry.incoming_bodies.get(&body_id).map(|b| b.body.clone());
    }
    if let Some(response_id) = resource_id(arg, KIND_INCOMING_RESPONSE) {
        let registry = registry().lock().unwrap();
        return registry
            .incoming_responses
            .get(&response_id)
            .map(|r| r.body.as_bytes().to_vec());
    }
    if let Some(response_id) = resource_id(arg, KIND_OUTGOING_RESPONSE) {
        let registry = registry().lock().unwrap();
        let body_id = registry.outgoing_responses.get(&response_id)?.body_id;
        return match body_id {
            Some(id) => registry.outgoing_bodies.get(&id).map(|b| b.bytes.clone()),
            // A response whose body was never opened has an empty body, which
            // is a legal answer — not a missing resource.
            None => Some(Vec::new()),
        };
    }
    if let Some(request_id) = resource_id(arg, KIND_INCOMING_REQUEST) {
        // §consume: "Will only return success at most once, and subsequent
        // calls will return error." The flag is shared with the 0.2
        // `incoming-request.consume`, so a program cannot take the body twice
        // by mixing the two spellings.
        let mut registry = registry().lock().unwrap();
        let request = registry.incoming_requests.get(&request_id)?;
        if request.consumed {
            return None;
        }
        let body = request.body.clone();
        if let Some(request) = registry.incoming_requests.get_mut(&request_id) {
            request.consumed = true;
        }
        return Some(body);
    }
    if let Some(request_id) = resource_id(arg, KIND_OUTGOING_REQUEST) {
        let registry = registry().lock().unwrap();
        // A request built with no `contents` has an EMPTY body, which is a
        // legal answer — not a missing resource.
        return Some(
            registry
                .outgoing_requests
                .get(&request_id)
                .map(|request| request.body.clone())
                .unwrap_or_default(),
        );
    }
    None
}

/// `tuple<stream<u8>, future<result<option<trailers>, error-code>>>` — the 0.3
/// body handover shape. The stream is closed before returning, so a guest
/// draining it with `canon stream.read` sees COMPLETED chunks then DROPPED,
/// and never BLOCKED.
fn body_stream_tuple(ctx: &mut HostContext, bytes: Vec<u8>) -> Value {
    let (stream_val, stream_id) = ctx.create_stream();
    for byte in &bytes {
        ctx.stream_push(stream_id, Value::I32(*byte as i32));
    }
    ctx.stream_close(stream_id);
    let (future_val, future_id) = ctx.create_future();
    ctx.resolve_future(future_id, Value::Null);
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        stream_val, future_val,
    ])))
}
