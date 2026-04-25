//! `vybe:crypto` — convenience hash primitives.
//!
//! Backed by audited RustCrypto crates (`sha2`, `md-5`). The previous
//! version inlined hand-rolled SHA-256 and MD5 with a "use a proper
//! impl for production" caveat — that's exactly the kind of ad-hoc
//! crypto the project's namespacing rules forbid. Real crates only.
//!
//! The vendored `wasi-crypto` proposal (`proposals/wasi-crypto/`,
//! Phase 2 in the WASI subgroup) is the standardized surface for
//! Wasm crypto, but it's a handle-based symmetric-state API
//! (`symmetric_state_open("SHA-256") → absorb → squeeze`) — not a
//! flat one-shot like `sha256(bytes) → hex`. The flat surface here
//! lives under `vybe:crypto` for the convenience-call sites; a real
//! `wasi:crypto` host module exposing the handle-based API is
//! separate work tracked under the namespace migration plan.

use std::sync::Arc;
use md5::{Digest as Md5Digest, Md5};
use sha2::{Digest as Sha2Digest, Sha256};
use vybe_bytecode::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:crypto", "sha256", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(sha256_hex(input.as_bytes()).as_str()))
    }));

    vm.register_host_fn("vybe:crypto", "md5", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(md5_hex(input.as_bytes()).as_str()))
    }));
}

pub(crate) fn sha256_hex(data: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(data);
    let bytes = hasher.finalize();
    bytes.iter().map(|byte| format!("{:02x}", byte)).collect()
}

pub(crate) fn md5_hex(data: &[u8]) -> String {
    let mut hasher = Md5::new();
    hasher.update(data);
    let bytes = hasher.finalize();
    bytes.iter().map(|byte| format!("{:02x}", byte)).collect()
}
