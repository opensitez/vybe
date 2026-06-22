//! `wasi:crypto/hashes` — digest convenience functions layered on the
//! `wasi-crypto` symmetric proposal.
//!
//! The full proposal at [`proposals/wasi-crypto/witx/witx-0.10/`] models
//! hashing as a sponge: `symmetric_state_open(algorithm)` →
//! `symmetric_state_absorb(state, data)` → `symmetric_state_squeeze(state, out)`.
//! That state-handle dance is overkill for the common
//! `sha256(bytes) -> hex` shape every Node/PHP/Ruby caller wants, so
//! Vybe registers a shim interface `wasi:crypto/hashes` that returns
//! the hex digest in one call. Higher-level wrappers (Node's
//! `crypto.createHash`, PHP `hash()`, etc.) compose this primitive.
//!
//! The hex output (rather than raw bytes) matches what hashing APIs
//! across host languages return by default and avoids an ArrayBuffer
//! round-trip for the typical "give me a hex string" use case.
//! Callers needing raw bytes use `web:crypto.digest` (WHATWG
//! SubtleCrypto) which returns an ArrayBuffer per spec.

use std::sync::Arc;
use vybe_bytecode::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:crypto/hashes",
        "sha256",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            Value::String(Arc::from(sha256_hex(input.as_bytes()).as_str()))
        }),
    );

    vm.register_host_fn(
        "wasi:crypto/hashes",
        "md5",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            Value::String(Arc::from(md5_hex(input.as_bytes()).as_str()))
        }),
    );

}

/// SHA-256 hex digest. Used directly by [`register`] above and reused
/// by Node-shaped wrappers in [`crate::node::crypto`].
pub fn sha256_hex(data: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    let result = Sha256::digest(data);
    result.iter().map(|b| format!("{:02x}", b)).collect()
}

/// MD5 hex digest. Same story as [`sha256_hex`].
pub fn md5_hex(data: &[u8]) -> String {
    use md5::{Digest, Md5};
    let result = Md5::digest(data);
    result.iter().map(|b| format!("{:02x}", b)).collect()
}
