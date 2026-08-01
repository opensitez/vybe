//! `node:http` header validation and module constants.
//!
//! `http.validateHeaderName(name)` and `http.validateHeaderValue(name, value)`
//! throw in Node rather than returning a flag; the VM has no host-fn exception
//! channel, so they answer with the WASI-style error object (`__wasi_error`)
//! the rest of the host layer uses, and `null` when valid.
//!
//! The rules are RFC 9110's, which is what Node's parser enforces:
//!   - a field name is a non-empty `token` (§5.6.2)
//!   - a field value carries no CR, LF or NUL, and does not lead or trail
//!     with whitespace (§5.5) — header injection is the reason this matters

use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

/// RFC 9110 §5.6.2 `tchar`.
fn is_token_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || "!#$%&'*+-.^_`|~".contains(c)
}

pub fn is_valid_header_name(name: &str) -> bool {
    !name.is_empty() && name.chars().all(is_token_char)
}

pub fn is_valid_header_value(value: &str) -> bool {
    if value.starts_with([' ', '\t']) || value.ends_with([' ', '\t']) {
        return false;
    }
    !value.contains(['\r', '\n', '\0'])
}

fn error(code: &str) -> Value {
    let mut object = Object::new();
    object
        .properties
        .insert("__wasi_error".into(), Value::String(Arc::from(code)));
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn string_arg(args: &[Value], index: usize) -> String {
    match args.get(index) {
        Some(Value::String(text)) => text.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:http",
        "validate_header_name",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if is_valid_header_name(&string_arg(args, 0)) {
                Value::Null
            } else {
                error("ERR_INVALID_HTTP_TOKEN")
            }
        }),
    );

    vm.register_host_fn(
        "node:http",
        "validate_header_value",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if is_valid_header_value(&string_arg(args, 1)) {
                Value::Null
            } else {
                error("ERR_INVALID_CHAR")
            }
        }),
    );

    // `http.maxHeaderSize` — Node's default cap on request headers, 16 KiB.
    vm.register_host_fn(
        "node:http",
        "max_header_size",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(16384.0)),
    );
}
