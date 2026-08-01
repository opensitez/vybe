//! `node:http` — HTTP server host module.
//!
//! See `httpserver.md` at the repo root for the full architecture. Briefly:
//! every supported language exposes idiomatic web APIs (PHP superglobals,
//! Node-style `(req, res)`, WSGI, Rack, ASP.NET `HttpContext`) as thin
//! **adapters** on top of the primitives declared here. Parsing, formatting,
//! MIME detection, compression, content negotiation, etc. happen exactly
//! once — in Rust — and every language calls into them.
//!
//! This module is organized as:
//!
//! - [`context`]  — per-request state + thread-local + drop guard
//! - [`request`]  — `node:http.*` (raw + parsed accessors)
//! - [`response`] — `node:http.*` (status, headers, write, end)
//! - [`tables`]   — `node:http.STATUS_CODES` / `.METHODS`
//!
//! What belongs here is Node's `http` module and nothing else. Request
//! shaping — the CGI environment, cookies, query pairs, form bodies — is
//! `wasi:http` read through `vybe_compiler::primitives::http_*`, and PHP's
//! own spellings (`header()`, `http_response_code()`, `php_sapi_name()`)
//! live in the PHP emitter. A `node:http` function with no counterpart in
//! Node is a bug, not a convenience.

pub mod client;
pub mod context;
pub mod request;
pub mod response;
pub mod tables;
pub mod server;
pub mod validate;

pub use context::{
    RequestBodyReader, RequestContext, ResponseMessage, ResponseState, install_context,
    take_context, with_context,
};

use vybe_runtime::VM;

/// Register every `node:http` host function on the VM.
///
/// Called from `register_with_capabilities` when `Capability::HttpServer`
/// is granted. Safe to call when no server is running — the primitives
/// simply report "cli" mode and return null / no-op when no
/// `RequestContext` is installed.
pub fn register(vm: &mut VM) {
    request::register(vm);
    response::register(vm);
    client::register(vm);
    tables::register(vm);
    validate::register(vm);
    server::register(vm);
}
