//! # `wasi:*` host imports — real WASI 0.2.8 proposals.
//!
//! Implementations match the WIT definitions vendored under
//! `proposals/wasi-*/wit/`. Function names match the canonical-ABI
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
pub mod io;
pub mod sql;
pub mod plugin;
pub use plugin::Plugin;

pub mod sockets;
pub mod clock;
pub mod console;
pub mod env;
pub mod random;
pub mod fs;

use vybe_runtime::VM;

/// Register the always-on WASI modules. Capability gating happens
/// in [`crate::modules`].
pub fn register(vm: &mut VM) {
    crypto::register(vm);
    filesystem::register(vm);
    // io::register runs after filesystem so its [method] handlers take precedence,
    // giving unified file+socket+fd dispatch for all wasi:io/streams resources.
    io::register(vm);
}

/// VM hot-reset: clear host-global state that lives OUTSIDE the VM heap — open
/// SQL connections and OS sockets — so a reused VM starts each run clean. Call
/// from the reset path alongside `VM::reset_to`. GUI state is reset separately
/// via `GuiState::reset` since the runner owns that `Arc`. See
/// `vmhotresetplan.md`.
pub fn reset_host_globals() {
    sql::reset();
    sockets::reset();
}
