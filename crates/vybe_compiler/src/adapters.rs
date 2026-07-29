//! In-language adapter modules — the Phase 6 Layer 3.
//!
//! Adapters are only valid when they re-export a real underlying surface.
//! The older `node:http -> wasi:http/server` example was not honest: there is
//! no real `wasi:http/server` host module in the tree today, and `node:http`
//! already exists as a real host module. Registering an adapter there would
//! overwrite the real module record with a recursive fake.
//!
//! Until a real adapter target exists, adapter registration stays empty.

use vybe_runtime::VM;

/// Register every bundled adapter module against the VM.
///
/// Order matters: adapters that re-export from other adapters must be
/// registered after their sources. Today all adapters re-export
/// exclusively from Synthetic (`wasi:*` / `wasm:js-*` / `vybe:*`)
/// modules, so order within this function is arbitrary.
///
/// Adapters are JS source files compiled at registration time. A parse
/// or link error aborts setup — adapters ship with the binary so any
/// breakage is a build error, not a runtime error.
pub fn register_all(_vm: &mut VM) -> Result<(), String> {
    Ok(())
}
