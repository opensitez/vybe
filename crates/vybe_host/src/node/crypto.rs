//! `node:crypto` — Node.js built-in `crypto` module.
//!
//! Reference: <https://nodejs.org/api/crypto.html>.
//!
//! Phase 1 surface: convenience hash functions backed by the
//! `wasi:crypto/hashes` shim in [`crate::wasi::crypto`]. Real Node uses
//! a streaming `crypto.createHash('algo').update(data).digest('hex')`
//! shape — the streaming form is deferred until a Hash object class
//! lands.

use std::sync::Arc;
use vybe_bytecode::{VM, Value};

use crate::wasi::crypto::{md5_hex, sha256_hex};

pub fn register(vm: &mut VM) {
    // Vybe-shorthand `sha256(data)` / `md5(data)` — convenience top-
    // level functions that mirror what the previous `node:crypto`
    // adapter exported. Not in real Node, but kept under this name
    // for source compatibility.
    vm.register_host_fn("node:crypto", "sha256", Box::new(|_ctx, args| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(sha256_hex(input.as_bytes()).as_str()))
    }));

    vm.register_host_fn("node:crypto", "md5", Box::new(|_ctx, args| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(md5_hex(input.as_bytes()).as_str()))
    }));
}
