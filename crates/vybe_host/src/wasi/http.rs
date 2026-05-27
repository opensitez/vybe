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

use vybe_bytecode::typedef::TypeDef;
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

const KIND_HEADERS: &str = "headers";
const KIND_OUTGOING_REQUEST: &str = "outgoing-request";
const KIND_REQUEST_OPTIONS: &str = "request-options";
const KIND_INCOMING_RESPONSE: &str = "incoming-response";
const KIND_FUTURE_INCOMING_RESPONSE: &str = "future-incoming-response";

#[derive(Clone, Copy)]
struct HttpTypeIds {
    headers: usize,
    outgoing_request: usize,
    request_options: usize,
    incoming_response: usize,
    future_incoming_response: usize,
}

#[derive(Debug, Clone)]
struct HeadersResource {
    entries: Vec<(String, Vec<u8>)>,
}

#[derive(Debug, Clone)]
struct OutgoingRequestResource {
    headers_id: u32,
    method: String,
    path_with_query: Option<String>,
    scheme: Option<String>,
    authority: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct RequestOptionsResource;

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
    object.properties.insert("__wasi_kind".into(), Value::String(Arc::from(kind)));
    object.properties.insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(Arc::new(Mutex::new(object)))
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
    Value::Object(Arc::new(Mutex::new(object)))
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
    Value::Object(Arc::new(Mutex::new(Object::new_array(bytes))))
}

fn header_values_array(values: &[Vec<u8>]) -> Value {
    let arrays = values
        .iter()
        .map(|value| header_value_bytes(&String::from_utf8_lossy(value)))
        .collect::<Vec<_>>();
    Value::Object(Arc::new(Mutex::new(Object::new_array(arrays))))
}

fn header_entries_array(entries: &[(String, Vec<u8>)]) -> Value {
    let pairs = entries
        .iter()
        .map(|(name, value)| {
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                Value::String(Arc::from(name.as_str())),
                header_value_bytes(&String::from_utf8_lossy(value)),
            ]))))
        })
        .collect::<Vec<_>>();
    Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))))
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
        Some(index) => (&host_port[..index], host_port[index + 1..].parse::<u16>().unwrap_or(80)),
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
        method,
        path,
        host,
        content_length,
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
        Some(index) => (&raw_response[..index], raw_response[index + 4..].to_string()),
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

    Ok(HttpResponseData { status, headers, body })
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
        future_incoming_response: resource(vm, "future-incoming-response", "HttpFutureIncomingResponse"),
    }
}

fn register_types(vm: &mut VM, type_ids: HttpTypeIds) {
    vm.register_host_fn("wasi:http/types", "[constructor]fields", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        let mut registry = registry().lock().unwrap();
        let id = registry.alloc_id();
        registry.headers.insert(id, HeadersResource { entries: Vec::new() });
        make_resource(KIND_HEADERS, id, type_ids.headers)
    }));

    vm.register_host_fn("wasi:http/types", "[method]fields.entries", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else { return err("invalid-argument"); };
        let registry = registry().lock().unwrap();
        let Some(headers) = registry.headers.get(&headers_id) else { return err("invalid-argument"); };
        header_entries_array(&headers.entries)
    }));

    vm.register_host_fn("wasi:http/types", "[method]fields.has", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else { return err("invalid-argument"); };
        let Some(name) = string_arg(args, 1) else { return Value::Bool(false); };
        let registry = registry().lock().unwrap();
        let Some(headers) = registry.headers.get(&headers_id) else { return err("invalid-argument"); };
        let target = name.to_ascii_lowercase();
        Value::Bool(headers.entries.iter().any(|(entry_name, _)| entry_name.eq_ignore_ascii_case(&target)))
    }));

    vm.register_host_fn("wasi:http/types", "[method]fields.get", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else { return err("invalid-argument"); };
        let Some(name) = string_arg(args, 1) else {
            return Value::Object(Arc::new(Mutex::new(Object::new_array(Vec::new()))));
        };
        let registry = registry().lock().unwrap();
        let Some(headers) = registry.headers.get(&headers_id) else { return err("invalid-argument"); };
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
    }));

    vm.register_host_fn("wasi:http/types", "[constructor]outgoing-request", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let Some(headers_id) = resource_id(&args[0], KIND_HEADERS) else { return err("invalid-argument"); };
        let mut registry = registry().lock().unwrap();
        let id = registry.alloc_id();
        registry.outgoing_requests.insert(
            id,
            OutgoingRequestResource {
                headers_id,
                method: "GET".into(),
                path_with_query: None,
                scheme: None,
                authority: None,
            },
        );
        make_resource(KIND_OUTGOING_REQUEST, id, type_ids.outgoing_request)
    }));

    vm.register_host_fn("wasi:http/types", "[method]outgoing-request.set-method", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else { return err("invalid-argument"); };
        let Some(method) = string_arg(args, 1) else { return err("HTTP-request-method-invalid"); };
        let mut registry = registry().lock().unwrap();
        let Some(request) = registry.outgoing_requests.get_mut(&request_id) else { return err("invalid-argument"); };
        if method.trim().is_empty() {
            return err("HTTP-request-method-invalid");
        }
        request.method = method.trim().to_ascii_uppercase();
        Value::Null
    }));

    vm.register_host_fn("wasi:http/types", "[method]outgoing-request.set-path-with-query", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else { return err("invalid-argument"); };
        let mut registry = registry().lock().unwrap();
        let Some(request) = registry.outgoing_requests.get_mut(&request_id) else { return err("invalid-argument"); };
        request.path_with_query = string_arg(args, 1);
        Value::Null
    }));

    vm.register_host_fn("wasi:http/types", "[method]outgoing-request.set-scheme", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else { return err("invalid-argument"); };
        let mut registry = registry().lock().unwrap();
        let Some(request) = registry.outgoing_requests.get_mut(&request_id) else { return err("invalid-argument"); };
        request.scheme = string_arg(args, 1).map(|value| value.to_ascii_lowercase());
        Value::Null
    }));

    vm.register_host_fn("wasi:http/types", "[method]outgoing-request.set-authority", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else { return err("invalid-argument"); };
        let mut registry = registry().lock().unwrap();
        let Some(request) = registry.outgoing_requests.get_mut(&request_id) else { return err("invalid-argument"); };
        request.authority = string_arg(args, 1);
        Value::Null
    }));

    vm.register_host_fn("wasi:http/types", "[method]outgoing-request.headers", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let Some(request_id) = resource_id(&args[0], KIND_OUTGOING_REQUEST) else { return err("invalid-argument"); };
        let registry = registry().lock().unwrap();
        let Some(request) = registry.outgoing_requests.get(&request_id) else { return err("invalid-argument"); };
        make_resource(KIND_HEADERS, request.headers_id, type_ids.headers)
    }));

    vm.register_host_fn("wasi:http/types", "[constructor]request-options", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        let mut registry = registry().lock().unwrap();
        let id = registry.alloc_id();
        registry.request_options.insert(id, RequestOptionsResource);
        make_resource(KIND_REQUEST_OPTIONS, id, type_ids.request_options)
    }));

    vm.register_host_fn("wasi:http/types", "[method]incoming-response.status", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(response_id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) else { return err("invalid-argument"); };
        let registry = registry().lock().unwrap();
        let Some(response) = registry.incoming_responses.get(&response_id) else { return err("invalid-argument"); };
        Value::F64(response.status as f64)
    }));

    vm.register_host_fn("wasi:http/types", "[method]incoming-response.headers", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let Some(response_id) = resource_id(&args[0], KIND_INCOMING_RESPONSE) else { return err("invalid-argument"); };
        let registry = registry().lock().unwrap();
        let Some(response) = registry.incoming_responses.get(&response_id) else { return err("invalid-argument"); };
        let _ = response.body.len();
        make_resource(KIND_HEADERS, response.headers_id, type_ids.headers)
    }));

    vm.register_host_fn("wasi:http/types", "[method]future-incoming-response.get", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let Some(future_id) = resource_id(&args[0], KIND_FUTURE_INCOMING_RESPONSE) else { return err("invalid-argument"); };
        let mut registry = registry().lock().unwrap();
        let Some(future) = registry.future_incoming_responses.get_mut(&future_id) else { return err("invalid-argument"); };
        if future.consumed {
            return err("already-consumed");
        }
        future.consumed = true;
        if let Some(code) = future.error.as_deref() {
            return err(code);
        }
        let Some(response_id) = future.response_id else { return Value::Null; };
        make_resource(KIND_INCOMING_RESPONSE, response_id, type_ids.incoming_response)
    }));
}

fn register_outgoing_handler(vm: &mut VM, type_ids: HttpTypeIds) {
    vm.register_host_fn("wasi:http/outgoing-handler", "handle", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
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
        let Some(authority) = request.authority.filter(|authority| !authority.trim().is_empty()) else {
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
                    registry.headers.insert(headers_id, HeadersResource { entries: response.headers });

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
    }));
}

fn register_legacy_shim(vm: &mut VM) {
    vm.register_host_fn("wasi:http", "get", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let url = string_arg(args, 0).unwrap_or_default();
        match http_request("GET", &url, None) {
            Ok(response) => Value::String(Arc::from(response.body.as_str())),
            Err(error) => Value::String(Arc::from(format!("Error: {}", error))),
        }
    }));

    vm.register_host_fn("wasi:http", "post", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let url = string_arg(args, 0).unwrap_or_default();
        let body = string_arg(args, 1).unwrap_or_default();
        match http_request("POST", &url, Some(&body)) {
            Ok(response) => Value::String(Arc::from(response.body.as_str())),
            Err(error) => Value::String(Arc::from(format!("Error: {}", error))),
        }
    }));

    vm.register_host_fn("wasi:http", "fetch", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let url = string_arg(args, 0).unwrap_or_default();
        let method = string_arg(args, 1).unwrap_or_else(|| "GET".into());
        let body = string_arg(args, 2);

        match http_request(&method, &url, body.as_deref()) {
            Ok(response) => {
                let mut object = Object::new();
                object.properties.insert("status".into(), Value::F64(response.status as f64));
                object
                    .properties
                    .insert("body".into(), Value::String(Arc::from(response.body.as_str())));
                object
                    .properties
                    .insert("ok".into(), Value::Bool((200..300).contains(&response.status)));
                Value::Object(Arc::new(Mutex::new(object)))
            }
            Err(error) => {
                let mut object = Object::new();
                object.properties.insert("status".into(), Value::F64(0.0));
                object
                    .properties
                    .insert("body".into(), Value::String(Arc::from(error.as_str())));
                object.properties.insert("ok".into(), Value::Bool(false));
                Value::Object(Arc::new(Mutex::new(object)))
            }
        }
    }));
}

pub fn register(vm: &mut VM) {
    let type_ids = register_resource_types(vm);
    register_types(vm, type_ids);
    register_outgoing_handler(vm, type_ids);
    register_legacy_shim(vm);
}
