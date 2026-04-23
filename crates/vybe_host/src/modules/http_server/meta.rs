//! `vybe:http/meta.*` — SAPI introspection.
//!
//! `mode()` is the polyglot equivalent of PHP's `php_sapi_name()`. Returns
//! one of:
//! - `"cli"` — no request context installed (terminal run, test, etc.)
//! - `"server-request"` — currently executing inside a request handler
//! - `"server-owner"` — Phase 2 programmatic mode: top-level of a script
//!   that has called `server.listen` but is not currently processing a
//!   request. Not distinguished in Phase 1; treated as `"cli"` until
//!   programmatic mode lands.

use std::sync::Arc;
use vybe_bytecode::{VM, Value, HostContext};

use super::context::with_context;

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:http/meta", "mode", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        match with_context(|_c| ()) {
            Some(_) => Value::String(Arc::from("server-request")),
            None    => Value::String(Arc::from("cli")),
        }
    }));

    vm.register_host_fn("vybe:http/meta", "server_software", Box::new(|_ctx, _| {
        Value::String(Arc::from(concat!("vybex/", env!("CARGO_PKG_VERSION"))))
    }));

    vm.register_host_fn("vybe:http/meta", "request_id", Box::new(|_ctx, _| {
        with_context(|c| Value::String(Arc::from(c.request_id.as_str())))
            .unwrap_or_else(|| Value::String(Arc::from("")))
    }));

    // PHP-idiom: `php_sapi_name()` returns "vybex-server" under --serve,
    // "cli" at the terminal. Matches real PHP's per-SAPI naming
    // convention (PHP's CLI SAPI returns "cli"; the built-in dev server
    // returns "cli-server").
    vm.register_host_fn("vybe:http/meta", "php_sapi_name", Box::new(|_ctx, _| {
        match with_context(|_c| ()) {
            Some(_) => Value::String(Arc::from("vybex-server")),
            None    => Value::String(Arc::from("cli")),
        }
    }));
}
