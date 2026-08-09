//! `node:http.STATUS_CODES` and `node:http.METHODS`.
//!
//! Node exports both as plain data, not functions:
//!
//! - `http.STATUS_CODES` — status code → standard reason phrase, used to fill
//!   `ServerResponse.statusMessage` when the script does not set one.
//! - `http.METHODS` — the methods the parser accepts, sorted.
//!
//! Values are IANA-registered reason phrases (RFC 9110 §15 and the HTTP
//! Status Code Registry), which is what Node ships.

use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

/// `(code, reason)` exactly as Node's `STATUS_CODES` spells them.
const STATUS_CODES: &[(u16, &str)] = &[
    (100, "Continue"),
    (101, "Switching Protocols"),
    (102, "Processing"),
    (103, "Early Hints"),
    (200, "OK"),
    (201, "Created"),
    (202, "Accepted"),
    (203, "Non-Authoritative Information"),
    (204, "No Content"),
    (205, "Reset Content"),
    (206, "Partial Content"),
    (207, "Multi-Status"),
    (208, "Already Reported"),
    (226, "IM Used"),
    (300, "Multiple Choices"),
    (301, "Moved Permanently"),
    (302, "Found"),
    (303, "See Other"),
    (304, "Not Modified"),
    (305, "Use Proxy"),
    (307, "Temporary Redirect"),
    (308, "Permanent Redirect"),
    (400, "Bad Request"),
    (401, "Unauthorized"),
    (402, "Payment Required"),
    (403, "Forbidden"),
    (404, "Not Found"),
    (405, "Method Not Allowed"),
    (406, "Not Acceptable"),
    (407, "Proxy Authentication Required"),
    (408, "Request Timeout"),
    (409, "Conflict"),
    (410, "Gone"),
    (411, "Length Required"),
    (412, "Precondition Failed"),
    (413, "Payload Too Large"),
    (414, "URI Too Long"),
    (415, "Unsupported Media Type"),
    (416, "Range Not Satisfiable"),
    (417, "Expectation Failed"),
    (418, "I'm a Teapot"),
    (421, "Misdirected Request"),
    (422, "Unprocessable Entity"),
    (423, "Locked"),
    (424, "Failed Dependency"),
    (425, "Too Early"),
    (426, "Upgrade Required"),
    (428, "Precondition Required"),
    (429, "Too Many Requests"),
    (431, "Request Header Fields Too Large"),
    (451, "Unavailable For Legal Reasons"),
    (500, "Internal Server Error"),
    (501, "Not Implemented"),
    (502, "Bad Gateway"),
    (503, "Service Unavailable"),
    (504, "Gateway Timeout"),
    (505, "HTTP Version Not Supported"),
    (506, "Variant Also Negotiates"),
    (507, "Insufficient Storage"),
    (508, "Loop Detected"),
    (509, "Bandwidth Limit Exceeded"),
    (510, "Not Extended"),
    (511, "Network Authentication Required"),
];

/// `http.METHODS` — sorted, as Node reports it.
const METHODS: &[&str] = &[
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
    "QUERY",
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

/// The standard reason phrase for `code`, if there is one.
pub fn reason_phrase(code: u16) -> Option<&'static str> {
    STATUS_CODES
        .iter()
        .find(|(c, _)| *c == code)
        .map(|(_, reason)| *reason)
}

pub fn register(vm: &mut VM) {
    // `http.STATUS_CODES` — an object keyed by the code AS A STRING, which is
    // what a JS object with numeric keys gives you.
    vm.register_host_fn(
        "node:http",
        "status_codes",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut entries = indexmap::IndexMap::new();
            for (code, reason) in STATUS_CODES {
                entries.insert(
                    Value::String(Arc::from(code.to_string().as_str())),
                    Value::String(Arc::from(*reason)),
                );
            }
            let mut object = Object::new();
            object.kind = ObjectKind::Map(entries);
            Value::Object(vybe_runtime::heap::alloc(object))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "methods",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let elems: Vec<Value> = METHODS
                .iter()
                .map(|m| Value::String(Arc::from(*m)))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
        }),
    );
}
