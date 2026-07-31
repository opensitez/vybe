//! `wasi:http` — minimal real WASI HTTP host surface plus legacy shim.
//!
//! This file now exposes a standards-shaped outgoing client slice:
//! - `wasi:http/types`
//! - `wasi:http/outgoing-handler`
//!
//! The implementation is intentionally small: enough to construct an
//! `outgoing-request`, send it through `outgoing-handler.handle`, and read back
//! the response status and headers. Server-side `wasi:http/incoming-handler`
//! remains future work.
//!
//! Vybe also keeps the older flat `wasi:http.{get,post,fetch}` helpers as a
//! compatibility shim for existing JS code.

use std::collections::HashMap;
use std::io::{Read, Write};
use std::sync::{Arc, Mutex, OnceLock};

use vybe_runtime::typedef::TypeDef;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

const KIND_HEADERS: &str = "headers";
const KIND_OUTGOING_REQUEST: &str = "outgoing-request";
const KIND_REQUEST_OPTIONS: &str = "request-options";
const KIND_INCOMING_RESPONSE: &str = "incoming-response";
const KIND_FUTURE_INCOMING_RESPONSE: &str = "future-incoming-response";
const KIND_INCOMING_BODY: &str = "incoming-body";
const KIND_FUTURE_TRAILERS: &str = "future-trailers";
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
    incoming_request: usize,
    response_outparam: usize,
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

#[derive(Debug, Clone)]
struct FutureIncomingResponseResource {
    response_id: Option<u32>,
    error: Option<String>,
    consumed: bool,
}

struct Registry {
    headers: HashMap<u32, HeadersResource>,
    outgoing_requests: HashMap<u32, OutgoingRequestResource>,
    request_options: HashMap<u32, RequestOptionsResource>,
    incoming_responses: HashMap<u32, IncomingResponseResource>,
    future_incoming_responses: HashMap<u32, FutureIncomingResponseResource>,
    incoming_bodies: HashMap<u32, IncomingBodyResource>,
    outgoing_responses: HashMap<u32, OutgoingResponseResource>,
    outgoing_bodies: HashMap<u32, OutgoingBodyResource>,
    incoming_requests: HashMap<u32, IncomingRequestResource>,
    response_outparams: HashMap<u32, ResponseOutparamResource>,
    next_id: u32,
}

impl Registry {
    fn new() -> Self {
        Self {
            headers: HashMap::new(),
            outgoing_requests: HashMap::new(),
            request_options: HashMap::new(),
            incoming_responses: HashMap::new(),
            future_incoming_responses: HashMap::new(),
            incoming_bodies: HashMap::new(),
            outgoing_responses: HashMap::new(),
            outgoing_bodies: HashMap::new(),
            incoming_requests: HashMap::new(),
            response_outparams: HashMap::new(),
            next_id: 1,
        }
    }

    fn alloc_id(&mut self) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        id
    }
}

fn registry() -> &'static Mutex<Registry> {
    static REGISTRY: OnceLock<Mutex<Registry>> = OnceLock::new();
    REGISTRY.get_or_init(|| Mutex::new(Registry::new()))
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
        incoming_request: resource(vm, "incoming-request", "HttpIncomingRequest"),
        response_outparam: resource(vm, "response-outparam", "HttpResponseOutparam"),
    }
}

fn register_types(vm: &mut VM, type_ids: HttpTypeIds) {
    vm.register_host_fn(
        "wasi:http/types",
        "[constructor]fields",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            let mut registry = registry().lock().unwrap();
            let id = registry.alloc_id();
            registry.headers.insert(
                id,
                HeadersResource {
                    entries: Vec::new(),
                },
            );
            make_resource(KIND_HEADERS, id, type_ids.headers)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]fields.entries",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(headers) = registry.headers.get(&headers_id) else {
                return err("invalid-argument");
            };
            header_entries_array(&headers.entries)
        }),
    );

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
        "[constructor]outgoing-request",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let id = registry.alloc_id();
            registry.outgoing_requests.insert(
                id,
                OutgoingRequestResource {
                    options_id: None,
                    headers_id,
                    method: "GET".into(),
                    path_with_query: None,
                    scheme: None,
                    authority: None,
                },
            );
            make_resource(KIND_OUTGOING_REQUEST, id, type_ids.outgoing_request)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.set-method",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let Some(method) = string_arg(args, 1) else {
                return err("HTTP-request-method-invalid");
            };
            let mut registry = registry().lock().unwrap();
            let Some(request) = registry.outgoing_requests.get_mut(&request_id) else {
                return err("invalid-argument");
            };
            if method.trim().is_empty() {
                return err("HTTP-request-method-invalid");
            }
            request.method = method.trim().to_ascii_uppercase();
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.set-path-with-query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(request) = registry.outgoing_requests.get_mut(&request_id) else {
                return err("invalid-argument");
            };
            request.path_with_query = string_arg(args, 1);
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.set-scheme",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(request) = registry.outgoing_requests.get_mut(&request_id) else {
                return err("invalid-argument");
            };
            request.scheme = string_arg(args, 1).map(|value| value.to_ascii_lowercase());
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.set-authority",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(request) = registry.outgoing_requests.get_mut(&request_id) else {
                return err("invalid-argument");
            };
            request.authority = string_arg(args, 1);
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.headers",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(request) = registry.outgoing_requests.get(&request_id) else {
                return err("invalid-argument");
            };
            make_resource(KIND_HEADERS, request.headers_id, type_ids.headers)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[constructor]request-options",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            let mut registry = registry().lock().unwrap();
            let id = registry.alloc_id();
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

    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-response.status",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(response_id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(response) = registry.incoming_responses.get(&response_id) else {
                return err("invalid-argument");
            };
            Value::F64(response.status as f64)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-response.headers",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(response_id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(response) = registry.incoming_responses.get(&response_id) else {
                return err("invalid-argument");
            };
            let _ = response.body.len();
            make_resource(KIND_HEADERS, response.headers_id, type_ids.headers)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]future-incoming-response.get",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(future_id) = resource_id(&args[0], KIND_FUTURE_INCOMING_RESPONSE) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(future) = registry.future_incoming_responses.get_mut(&future_id) else {
                return err("invalid-argument");
            };
            if future.consumed {
                return err("already-consumed");
            }
            future.consumed = true;
            if let Some(code) = future.error.as_deref() {
                return err(code);
            }
            let Some(response_id) = future.response_id else {
                return Value::Null;
            };
            make_resource(
                KIND_INCOMING_RESPONSE,
                response_id,
                type_ids.incoming_response,
            )
        }),
    );

    // future-incoming-response.subscribe → pollable (always ready in sync model)
    vm.register_host_fn(
        "wasi:http/types",
        "[method]future-incoming-response.subscribe",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Pollable")));
            obj.properties.insert("__ready".into(), Value::Bool(true));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

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
            headers
                .entries
                .retain(|(k, _)| !k.eq_ignore_ascii_case(&key));
            if let Some(Value::Object(arr)) = args.get(2) {
                let inner = arr.lock().unwrap();
                if let vybe_runtime::value::ObjectKind::Array(ref elems) = inner.kind {
                    for val in elems {
                        let bytes = format!("{}", val).into_bytes();
                        headers.entries.push((key.clone(), bytes));
                    }
                }
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
            let new_id = registry.alloc_id();
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
            let id = registry.alloc_id();
            registry.headers.insert(id, HeadersResource { entries });
            make_resource(KIND_HEADERS, id, type_ids.headers)
        }),
    );

    // outgoing-request getters
    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.method",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(req) = registry.outgoing_requests.get(&id) else {
                return err("invalid-argument");
            };
            Value::String(Arc::from(req.method.as_str()))
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.path-with-query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(req) = registry.outgoing_requests.get(&id) else {
                return err("invalid-argument");
            };
            req.path_with_query
                .as_deref()
                .map(|s| Value::String(Arc::from(s)))
                .unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.scheme",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(req) = registry.outgoing_requests.get(&id) else {
                return err("invalid-argument");
            };
            req.scheme
                .as_deref()
                .map(|s| Value::String(Arc::from(s)))
                .unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.authority",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(req) = registry.outgoing_requests.get(&id) else {
                return err("invalid-argument");
            };
            req.authority
                .as_deref()
                .map(|s| Value::String(Arc::from(s)))
                .unwrap_or(Value::Null)
        }),
    );

    // [method]outgoing-request.body → result<outgoing-body, error-code>
    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-request.body",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(_request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let body_id = registry.alloc_id();
            registry
                .outgoing_bodies
                .insert(body_id, OutgoingBodyResource { bytes: Vec::new(), finished: false, trailers: Vec::new() });
            make_resource(KIND_OUTGOING_BODY, body_id, type_ids.outgoing_body)
        }),
    );

    // request-options timeout getters/setters (durations in nanoseconds)
    vm.register_host_fn(
        "wasi:http/types",
        "[method]request-options.connect-timeout",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(opt) = registry.request_options.get(&id) else {
                return err("invalid-argument");
            };
            opt.connect_timeout_ns
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null)
        }),
    );

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
        "[method]request-options.first-byte-timeout",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(opt) = registry.request_options.get(&id) else {
                return err("invalid-argument");
            };
            opt.first_byte_timeout_ns
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null)
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
        "[method]request-options.between-bytes-timeout",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_REQUEST_OPTIONS) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(opt) = registry.request_options.get(&id) else {
                return err("invalid-argument");
            };
            opt.between_bytes_timeout_ns
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null)
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
    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-response.consume",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(response_id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(response) = registry.incoming_responses.get(&response_id) else {
                return err("invalid-argument");
            };
            let body_bytes = response.body.as_bytes().to_vec();
            let body_id = registry.alloc_id();
            registry.incoming_bodies.insert(
                body_id,
                IncomingBodyResource {
                    body: body_bytes,
                    position: 0,
                },
            );
            make_resource(KIND_INCOMING_BODY, body_id, type_ids.incoming_body)
        }),
    );

    // incoming-body.%stream → input-stream resource readable via [method]input-stream.blocking-read
    // Registers the body bytes as a Buffer stream in the filesystem registry so standard
    // stream host functions can drain it without any bespoke logic in http.rs.
    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-body.%stream",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(body_id) = resource_id(&args[0], KIND_INCOMING_BODY) else {
                return err("invalid-argument");
            };
            let body_bytes = {
                let registry = registry().lock().unwrap();
                let Some(body) = registry.incoming_bodies.get(&body_id) else {
                    return err("invalid-argument");
                };
                body.body.clone()
            };
            let stream_id = super::filesystem::register_buffer_stream(body_bytes);
            let mut obj = Object::new();
            obj.properties.insert(
                "__wasi_kind".into(),
                Value::String(Arc::from("input-stream")),
            );
            obj.properties
                .insert("__wasi_id".into(), Value::F64(stream_id as f64));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // [static]incoming-body.finish → future-trailers
    vm.register_host_fn(
        "wasi:http/types",
        "[static]incoming-body.finish",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            let mut registry = registry().lock().unwrap();
            let id = registry.alloc_id();
            make_resource(KIND_FUTURE_TRAILERS, id, type_ids.future_trailers)
        }),
    );

    // future-trailers.get → option<result<option<trailers>, error-code>>
    vm.register_host_fn(
        "wasi:http/types",
        "[method]future-trailers.get",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            // No trailers in our HTTP/1.1 client — return Some(Ok(None))
            Value::Null
        }),
    );

    // future-trailers.subscribe → always-ready pollable
    vm.register_host_fn(
        "wasi:http/types",
        "[method]future-trailers.subscribe",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Pollable")));
            obj.properties.insert("__ready".into(), Value::Bool(true));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // [constructor]outgoing-response(fields) → outgoing-response
    vm.register_host_fn(
        "wasi:http/types",
        "[constructor]outgoing-response",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let headers_id = resource_id(&args[0], KIND_HEADERS).unwrap_or(0);
            let mut registry = registry().lock().unwrap();
            let id = registry.alloc_id();
            registry.outgoing_responses.insert(
                id,
                OutgoingResponseResource {
                    status: 200,
                    headers_id,
                    body_id: None,
                },
            );
            make_resource(KIND_OUTGOING_RESPONSE, id, type_ids.outgoing_response)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-response.status-code",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(resp) = registry.outgoing_responses.get(&id) else {
                return err("invalid-argument");
            };
            Value::F64(resp.status as f64)
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-response.set-status-code",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) else {
                return err("invalid-argument");
            };
            // §outgoing-response.set-status-code: "Fails if the status-code
            // given is not a valid http status code." Same rule the 0.3
            // `response.set-status-code` already enforced; this arm accepted
            // anything, so 999 silently became the response status.
            let code = args.get(1).map(|v| v.as_f64() as i64).unwrap_or(200);
            if !(100..=599).contains(&code) {
                return err("invalid-argument");
            }
            let mut registry = registry().lock().unwrap();
            let Some(resp) = registry.outgoing_responses.get_mut(&id) else {
                return err("invalid-argument");
            };
            resp.status = code as u16;
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-response.headers",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) else {
                return err("invalid-argument");
            };
            let registry = registry().lock().unwrap();
            let Some(resp) = registry.outgoing_responses.get(&id) else {
                return err("invalid-argument");
            };
            make_resource(KIND_HEADERS, resp.headers_id, type_ids.headers)
        }),
    );

    // [method]outgoing-response.body → result<outgoing-body, error-code>
    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-response.body",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(id) = resource_id(&args[0], KIND_OUTGOING_RESPONSE) else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let body_id = registry.alloc_id();
            registry
                .outgoing_bodies
                .insert(body_id, OutgoingBodyResource { bytes: Vec::new(), finished: false, trailers: Vec::new() });
            if let Some(resp) = registry.outgoing_responses.get_mut(&id) {
                resp.body_id = Some(body_id);
            }
            make_resource(KIND_OUTGOING_BODY, body_id, type_ids.outgoing_body)
        }),
    );

    // [method]outgoing-body.write → result<output-stream, error-code>
    // Returns a simplified stream object with __body_id for later flush.
    vm.register_host_fn(
        "wasi:http/types",
        "[method]outgoing-body.write",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(body_id) = resource_id(&args[0], KIND_OUTGOING_BODY) else {
                return err("invalid-argument");
            };
            let bytes_val = args.get(1).cloned().unwrap_or(Value::Null);
            if let Value::Object(arr) = &bytes_val {
                let inner = arr.lock().unwrap();
                if let vybe_runtime::value::ObjectKind::Array(ref elems) = inner.kind {
                    let bytes: Vec<u8> = elems.iter().map(|v| v.as_f64() as u8).collect();
                    drop(inner);
                    let mut registry = registry().lock().unwrap();
                    if let Some(body) = registry.outgoing_bodies.get_mut(&body_id) {
                        body.bytes.extend_from_slice(&bytes);
                    }
                }
            }
            Value::F64(0.0) // returns number of bytes written (simplified)
        }),
    );

    // [static]outgoing-body.finish → result<_, error-code>
    vm.register_host_fn(
        "wasi:http/types",
        "[static]outgoing-body.finish",
        // §outgoing-body.finish(this, trailers) -> result<_, error-code>.
        // "This must be called to signal that the response is complete."
        // Finishing twice is an error, matching the write-at-most-once rule on
        // the same resource.
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(body_id) = args.first().and_then(|a| resource_id(a, KIND_OUTGOING_BODY))
            else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(body) = registry.outgoing_bodies.get_mut(&body_id) else {
                return err("invalid-argument");
            };
            if body.finished {
                return err("invalid-argument");
            }
            body.finished = true;
            // `trailers` is `option<trailers>`; when present it is a headers
            // resource whose entries are appended to the body's trailer set.
            if let Some(trailers_id) = args.get(1).and_then(|a| resource_id(a, KIND_HEADERS)) {
                let entries = registry
                    .headers
                    .get(&trailers_id)
                    .map(|h| h.entries.clone())
                    .unwrap_or_default();
                if let Some(body) = registry.outgoing_bodies.get_mut(&body_id) {
                    body.trailers = entries;
                }
            }
            Value::Null
        }),
    );

    // http-error-code(err) → option<error-code>
    vm.register_host_fn(
        "wasi:http/types",
        "http-error-code",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let inner = obj.lock().unwrap();
                if let Some(code) = inner.properties.get("__wasi_error") {
                    return code.clone();
                }
            }
            Value::Null
        }),
    );

    // ── incoming-request (server side) ─────────────────────────────────────
    //
    // wasi:http/types §incoming-request. The host builds one of these from the
    // transport it is serving via `push_incoming_request`; guest code reads it
    // through these accessors. Previously all six returned `Value::Null` as
    // link padding.
    fn with_incoming_request<T>(
        args: &[Value],
        f: impl FnOnce(&IncomingRequestResource) -> T,
    ) -> Option<T> {
        let id = resource_id(args.first()?, KIND_INCOMING_REQUEST)?;
        let registry = registry().lock().unwrap();
        registry.incoming_requests.get(&id).map(f)
    }

    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-request.method",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match with_incoming_request(args, |r| r.method.clone()) {
                Some(m) => Value::String(Arc::from(m.as_str())),
                None => err("invalid-argument"),
            }
        }),
    );

    // `path-with-query`, `scheme` and `authority` are `option<...>` in the WIT:
    // absent is null, not an error.
    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-request.path-with-query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match with_incoming_request(args, |r| r.path_with_query.clone()) {
                Some(Some(p)) => Value::String(Arc::from(p.as_str())),
                Some(None) => Value::Null,
                None => err("invalid-argument"),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-request.scheme",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match with_incoming_request(args, |r| r.scheme.clone()) {
                Some(Some(v)) => Value::String(Arc::from(v.as_str())),
                Some(None) => Value::Null,
                None => err("invalid-argument"),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-request.authority",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match with_incoming_request(args, |r| r.authority.clone()) {
                Some(Some(v)) => Value::String(Arc::from(v.as_str())),
                Some(None) => Value::Null,
                None => err("invalid-argument"),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-request.headers",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            match with_incoming_request(args, |r| r.headers_id) {
                Some(headers_id) => make_resource(KIND_HEADERS, headers_id, type_ids.headers),
                None => err("invalid-argument"),
            }
        }),
    );

    // §incoming-request.consume: succeeds at most ONCE; later calls are errors.
    vm.register_host_fn(
        "wasi:http/types",
        "[method]incoming-request.consume",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = args.first().and_then(|a| resource_id(a, KIND_INCOMING_REQUEST))
            else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(request) = registry.incoming_requests.get(&request_id) else {
                return err("invalid-argument");
            };
            if request.consumed {
                return err("invalid-argument");
            }
            let body = request.body.clone();
            if let Some(request) = registry.incoming_requests.get_mut(&request_id) {
                request.consumed = true;
            }
            let body_id = registry.alloc_id();
            registry.incoming_bodies.insert(
                body_id,
                IncomingBodyResource { body, position: 0 },
            );
            make_resource(KIND_INCOMING_BODY, body_id, type_ids.incoming_body)
        }),
    );

    // ── response-outparam (server side) ────────────────────────────────────
    //
    // §response-outparam.set: "Set the value of the `response-outparam` to
    // either send a response, or indicate an error. This method consumes the
    // `response-outparam` to ensure that it is called at most once."
    vm.register_host_fn(
        "wasi:http/types",
        "[static]response-outparam.set",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(param_id) = args.first().and_then(|a| resource_id(a, KIND_RESPONSE_OUTPARAM))
            else {
                return err("invalid-argument");
            };
            let mut registry = registry().lock().unwrap();
            let Some(param) = registry.response_outparams.get(&param_id) else {
                return err("invalid-argument");
            };
            if param.set {
                return err("invalid-argument");
            }
            // arg1 is `result<outgoing-response, error-code>`: a response
            // resource on success, anything else read as the error code.
            let response_id = args.get(1).and_then(|a| resource_id(a, KIND_OUTGOING_RESPONSE));
            let error = match (&response_id, args.get(1)) {
                (None, Some(Value::String(code))) => Some(code.to_string()),
                (None, _) => Some("internal-error".to_string()),
                _ => None,
            };
            if let Some(param) = registry.response_outparams.get_mut(&param_id) {
                param.response_id = response_id;
                param.error = error;
                param.set = true;
            }
            Value::Null
        }),
    );

    // §response-outparam.send-informational — 1xx interim responses.
    vm.register_host_fn(
        "wasi:http/types",
        "[method]response-outparam.send-informational",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(param_id) = args.first().and_then(|a| resource_id(a, KIND_RESPONSE_OUTPARAM))
            else {
                return err("invalid-argument");
            };
            let status = match args.get(1) {
                Some(Value::F64(n)) => *n as u16,
                Some(Value::I32(n)) => *n as u16,
                _ => return err("invalid-argument"),
            };
            // Only 1xx codes are informational (RFC 9110 §15.2).
            if !(100..200).contains(&status) {
                return err("invalid-argument");
            }
            let headers_id = args
                .get(2)
                .and_then(|a| resource_id(a, KIND_HEADERS))
                .unwrap_or(0);
            let mut registry = registry().lock().unwrap();
            let Some(param) = registry.response_outparams.get_mut(&param_id) else {
                return err("invalid-argument");
            };
            param.informational.push((status, headers_id));
            Value::Null
        }),
    );
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
    let headers_id = registry.alloc_id();
    registry
        .headers
        .insert(headers_id, HeadersResource { entries: headers });
    let id = registry.alloc_id();
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
    let id = registry.alloc_id();
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

/// Build the guest-facing resource values for a served request.
pub fn incoming_request_value(vm: &VM, request_id: u32) -> Option<Value> {
    let type_id = vm
        .type_registry
        .get_id("HttpIncomingRequest")?;
    Some(make_resource(KIND_INCOMING_REQUEST, request_id, type_id))
}

pub fn response_outparam_value(vm: &VM, param_id: u32) -> Option<Value> {
    let type_id = vm
        .type_registry
        .get_id("HttpResponseOutparam")?;
    Some(make_resource(KIND_RESPONSE_OUTPARAM, param_id, type_id))
}

/// `wasi:http/outgoing-handler` — send an `outgoing-request` and get back a
/// `future-incoming-response`.
fn register_outgoing_handler(vm: &mut VM, type_ids: HttpTypeIds) {
    vm.register_host_fn(
        "wasi:http/outgoing-handler",
        "handle",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else {
                return err("HTTP-request-denied");
            };
            if args.get(1).is_some() && !matches!(args.get(1), Some(Value::Null)) {
                let Some(_options_id) = resource_id(&args[1], KIND_REQUEST_OPTIONS) else {
                    return err("HTTP-request-denied");
                };
            }

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

            let future_id = {
                let mut registry = registry().lock().unwrap();
                let future_id = registry.alloc_id();
                match http_request(&request.method, &url, None) {
                    Ok(response) => {
                        let headers_id = registry.alloc_id();
                        registry.headers.insert(
                            headers_id,
                            HeadersResource {
                                entries: response.headers,
                            },
                        );

                        let response_id = registry.alloc_id();
                        registry.incoming_responses.insert(
                            response_id,
                            IncomingResponseResource {
                                status: response.status,
                                headers_id,
                                body: response.body,
                            },
                        );

                        registry.future_incoming_responses.insert(
                            future_id,
                            FutureIncomingResponseResource {
                                response_id: Some(response_id),
                                error: None,
                                consumed: false,
                            },
                        );
                    }
                    Err(message) => {
                        registry.future_incoming_responses.insert(
                            future_id,
                            FutureIncomingResponseResource {
                                response_id: None,
                                error: Some(map_transport_error(&message).into()),
                                consumed: false,
                            },
                        );
                    }
                }
                future_id
            };

            make_resource(
                KIND_FUTURE_INCOMING_RESPONSE,
                future_id,
                type_ids.future_incoming_response,
            )
        }),
    );
}

pub fn register(vm: &mut VM) {
    let type_ids = register_resource_types(vm);
    register_types(vm, type_ids);
    register_outgoing_handler(vm, type_ids);
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
    fn with_request<T>(args: &[Value], f: impl FnOnce(&OutgoingRequestResource) -> T) -> Option<T> {
        let request_id = resource_id(&args[0], KIND_OUTGOING_REQUEST)?;
        let registry = registry().lock().unwrap();
        registry.outgoing_requests.get(&request_id).map(f)
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
            (|r: &OutgoingRequestResource| r.path_with_query.clone()) as fn(&_) -> Option<String>,
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
            let new_id = registry.alloc_id();
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
        ("[method]request.set-scheme", |r, v| r.scheme = v),
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

/// WASI 0.3 (`wasi:http@0.3.0-rc-2025-09-16`) `client` + `handler` interfaces.
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
    for (module, name) in [("wasi:http/client", "send"), ("wasi:http/handler", "handle")] {
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
                        let headers_id = registry.alloc_id();
                        registry.headers.insert(
                            headers_id,
                            HeadersResource {
                                entries: response.headers,
                            },
                        );
                        let response_id = registry.alloc_id();
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
    // [static]request.new — WASI 0.3 constructor.
    // Contents stream accepted but not drained (HostContext::stream_drain not yet available).
    vm.register_host_fn(
        "wasi:http/types",
        "[static]request.new",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let headers_id = resource_id(&args[0], KIND_HEADERS).unwrap_or(0);
            // §request.new(headers, contents, trailers, options) — the 4th
            // argument is `option<request-options>`, read back by
            // `request.get-options`. Accept it in either the spec position or
            // immediately after headers, since callers that pass no body still
            // want options to land.
            let options_id = args
                .get(3)
                .and_then(|a| resource_id(a, KIND_REQUEST_OPTIONS))
                .or_else(|| args.get(1).and_then(|a| resource_id(a, KIND_REQUEST_OPTIONS)));
            let mut reg = registry().lock().unwrap();
            let id = reg.alloc_id();
            reg.outgoing_requests.insert(
                id,
                OutgoingRequestResource {
                    options_id,
                    headers_id,
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
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let bytes = if let Some(body_id) = resource_id(&args[0], KIND_OUTGOING_BODY) {
                registry()
                    .lock()
                    .unwrap()
                    .outgoing_bodies
                    .get(&body_id)
                    .map(|b| b.bytes.clone())
                    .unwrap_or_default()
            } else {
                Vec::new()
            };
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
        }),
    );

    // [static]response.new — WASI 0.3 constructor alongside [constructor]outgoing-response.
    vm.register_host_fn(
        "wasi:http/types",
        "[static]response.new",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let headers_id = resource_id(&args[0], KIND_HEADERS).unwrap_or(0);
            let mut reg = registry().lock().unwrap();
            let id = reg.alloc_id();
            reg.outgoing_responses.insert(
                id,
                OutgoingResponseResource {
                    status: 200,
                    headers_id,
                    body_id: None,
                },
            );
            make_resource(KIND_OUTGOING_RESPONSE, id, type_ids.outgoing_response)
        }),
    );

    // [static]response.consume-body — returns body bytes as stream<u8> + resolved trailers future.
    vm.register_host_fn(
        "wasi:http/types",
        "[static]response.consume-body",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let bytes = if let Some(body_id) = resource_id(&args[0], KIND_OUTGOING_BODY) {
                registry()
                    .lock()
                    .unwrap()
                    .outgoing_bodies
                    .get(&body_id)
                    .map(|b| b.bytes.clone())
                    .unwrap_or_default()
            } else if let Some(body_id) = resource_id(&args[0], KIND_INCOMING_BODY) {
                registry()
                    .lock()
                    .unwrap()
                    .incoming_bodies
                    .get(&body_id)
                    .map(|b| b.body.clone())
                    .unwrap_or_default()
            } else {
                Vec::new()
            };
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
        }),
    );
}
