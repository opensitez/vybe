//! # `wasi:*` host imports — real WASI 0.3.1 proposals.
//!
//! Implementations match the WIT definitions vendored under
//! `proposals/WASI/proposals/*/wit/`, every one of which declares
//! `@0.3.1`. Read those, not this header: the versions claimed in
//! this directory were stale by four minor revisions until
//! 2026-08-21, and `interface_coverage.rs` is the only statement of
//! the surface that a test can check. Function names match the canonical-ABI
//! shape (`[method]<resource>.<name>`) so an external Component
//! Model runtime that loads Vybe-emitted `.wasm` can satisfy the
//! imports against any conforming WASI implementation.
//!
//! Only modules listed here use the `wasi:` namespace honestly. The
//! older Vybe-shim filesystem surface (flat-string-path APIs that
//! look nothing like WASI) lives under `vybe:fs`.

pub mod crypto;
pub mod filesystem;
pub mod http;
pub mod plugin;
pub mod sql;
pub use plugin::Plugin;

pub mod clock;
pub mod console;
pub mod env;
pub mod random;
pub mod sockets;
pub mod tls;

use vybe_runtime::VM;

/// Register the always-on WASI modules. Capability gating happens
/// in [`crate::modules`].
pub fn register(vm: &mut VM) {
    crypto::register(vm);
    filesystem::register(vm);
    tls::register(vm);
}

/// VM hot-reset: the state this platform still clears BY HAND, because clearing
/// it is an action and not just a drop — live SQL connections and OS sockets
/// hold kernel handles that want closing, and `sockets` shares its table with
/// spawned threads, so it cannot live in the thread-local resource store.
///
/// Everything else this platform holds for a running program — descriptors and
/// directory cursors, in-flight HTTP resources, key material — is VM-owned
/// storage ([`vybe_runtime::resources`]) and is dropped by `VM::reset_to`
/// itself. Those three used to be missing here: this function covered two of
/// five tables, so a reused VM handed the next program the previous one's
/// descriptors, response bodies and crypto keys. Storage the VM owns cannot be
/// left out of a list, because there is no list.
///
/// Called by the wasi plugin's `reset`, which `VM::reset_to` runs — no embedder
/// needs to call it.
pub fn reset_host_globals() {
    sql::reset();
    sockets::reset();
}
