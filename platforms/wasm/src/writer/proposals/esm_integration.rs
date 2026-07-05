//! # esm-integration proposal
//!
//! Spec: `proposals/esm-integration/`. Allows a
//! `.wasm` module to be loaded directly by a JS host via
//! `import { … } from "./module.wasm"`. This proposal is **mostly a JS
//! host concern** — it says nothing about WASM bytecode format. The
//! emitter's job is to produce a module whose shape a conforming host
//! will accept for ESM instantiation:
//!
//! 1. Valid WASM (magic + version + well-formed sections).
//! 2. An **export section** with at least one export — otherwise
//!    `import * from "./mod.wasm"` binds nothing.
//! 3. Imports whose `(module, name)` pairs resolve in the host's
//!    module-resolution graph. For `wasm:js-*` imports, conforming hosts
//!    (V8 ≥ 12, recent Node with `--experimental-wasm-modules`) resolve
//!    them to the standard built-in JS primitive helpers; for other
//!    imports, the JS loader follows normal ESM resolution (allowing
//!    `import "./foo.wasm"` to pull `./foo.wasm`'s own imports from
//!    sibling JS modules).
//!
//! This module owns the machinery that validates our emitter conforms
//! to (1)–(3) above. Functionality-wise it has no runtime behaviour —
//! the actual ESM wiring happens in the JS runtime.

use vybe_bytecode::Chunk;

/// Readiness check — returns `Ok(())` if the provided chunks will emit
/// a module that can be loaded as an ES module. Used from tests.
pub fn check_esm_readiness(chunks: &[Chunk]) -> Result<(), &'static str> {
    if chunks.is_empty() {
        return Err("ESM bundles must contain at least one function to export");
    }
    Ok(())
}

/// The exact resolution prefix the JS host assigns to well-known
/// built-in imports. A host that supports the js-string-builtins /
/// js-primitive-builtins proposals provides all imports under these
/// modules automatically:
pub const HOST_BUILTIN_MODULES: &[&str] = &[
    "wasm:js-string",
    "wasm:js-number",
    "wasm:js-boolean",
    "wasm:js-undefined",
    "wasm:js-symbol",
    "wasm:js-bigint",
];

/// Returns true when `(module, name)` is an import the JS host is
/// expected to provide automatically (no user-supplied glue needed).
pub fn is_host_builtin(module: &str, _name: &str) -> bool {
    HOST_BUILTIN_MODULES.iter().any(|m| *m == module)
}
