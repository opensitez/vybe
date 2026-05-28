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
//! - [`meta`]     — `node:http.*` (mode, server_software, request_id)
//!
//! Additional utility sub-modules (url, cookie, form, mime, date, etag,
//! range, negotiate, compress, auth, ws, sse, session, static_serve,
//! server) will land in subsequent slices.

pub mod context;
pub mod request;
pub mod response;
pub mod meta;
pub mod client;

pub use context::{
    RequestContext, RequestBodyReader, ResponseState, ResponseMessage,
    install_context, take_context, with_context,
};

use vybe_bytecode::VM;

/// Register every `node:http` host function on the VM.
///
/// Called from `register_with_capabilities` when `Capability::HttpServer`
/// is granted. Safe to call when no server is running — the primitives
/// simply report "cli" mode and return null / no-op when no
/// `RequestContext` is installed.
pub fn register(vm: &mut VM) {
    request::register(vm);
    response::register(vm);
    meta::register(vm);
    client::register(vm);
}
