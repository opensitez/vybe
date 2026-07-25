pub mod builtin_types;
pub mod controls;
pub use vybe_platform_ecma as ecma; // ecma:*    — ECMA-262 JS runtime (extracted crate)
pub mod modules;
pub mod namespaces;
pub use vybe_platform_node as node; // node:*    — Node.js built-in modules (extracted crate)
pub mod stdlib_aliases;
pub use vybe_platform_wasi as wasi; // wasi:*    — real WASI 0.2.8 proposals (extracted crate; incl. sockets)
pub mod wasm; // wasm:*    — real WebAssembly CG proposals
pub use vybe_platform_web as web; // web:*     — WHATWG / W3C web platform APIs (extracted crate)

#[cfg(feature = "gui")]
pub mod gui_state;

pub use controls::get_def;
#[cfg(feature = "gui")]
pub use gui_state::GuiState;
pub use modules::{Capabilities, Capability, register_all, register_with_capabilities};
#[cfg(feature = "gui")]
pub use modules::{register_all_with_gui, register_with_capabilities_and_gui};
pub use namespaces::setup_namespaces;
pub use node::http::{
    RequestBodyReader, RequestContext, ResponseMessage, ResponseState, install_context,
    take_context, with_context,
};
pub use wasi::sql::{fetch_columns_for_query, test_connection_and_list_tables};

/// VM hot-reset (bucket C): force-init every process-global shared ecma
/// prototype (and `globalThis`) so each is captured in the baseline a
/// `VM::snapshot` records. Call ONCE on the boot thread, after
/// `vybe_bytecode::heap::enable_tracking()` and before `snapshot()`. Without
/// this, a prototype first touched by a script AFTER the snapshot would be
/// wiped by a later reset. See `vmhotresetplan.md`.
pub use ecma::prime_shared_prototypes;

/// VM hot-reset (bucket C/D): clear host-global state that lives OUTSIDE the VM
/// heap — open SQL connections and OS sockets — so a reused VM starts each run
/// clean. Call from the reset path alongside `VM::reset_to`. GUI state is reset
/// separately via [`GuiState::reset`] since the runner owns that `Arc`.
/// See `vmhotresetplan.md`.
pub fn reset_host_globals() {
    wasi::sql::reset();
    wasi::sockets::reset();
}
