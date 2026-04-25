//! # `node:*` host imports — Node.js built-in modules.
//!
//! Node.js publishes its built-in modules under the `node:` import
//! prefix (`node:fs`, `node:path`, `node:os`, `node:crypto`, …) — a
//! convention adopted by Deno and Bun for Node compat. The Node fs/os
//! API is a *de facto* standard (not standardised by ECMA-262, WASI,
//! or WinterCG) and is documented at <https://nodejs.org/api/fs.html>
//! and <https://nodejs.org/api/os.html>.
//!
//! These host fns mirror the Node API faithfully — function names
//! match (`readFileSync`, `writeFileSync`, `existsSync`, …) and
//! return shapes match (Stats objects with `isFile()`/`isDirectory()`
//! methods, `mtimeMs` as a number, etc.).
//!
//! Compare the namespace landscape:
//!
//! - `wasm:js-*` — real WebAssembly CG proposals (only `wasm:js-string`
//!   merged; primitive builtins Stage-1).
//! - `wasi:*` — real WASI proposals. `wasi:filesystem` (descriptor-
//!   based) is the standardized filesystem API; this module is **not**
//!   that — it's the Node convention which most JS code expects.
//! - `ecma:*` — ECMA-262 spec types (Math, Array, JSON, …). ECMA does
//!   not specify I/O; fs/os live here under `node:*` instead.
//! - `vybe:*` — Vybe-only convenience (GUI, debug, cross-language).

pub mod child_process;
pub mod crypto;
pub mod fs;
pub mod os;
pub mod path;
pub mod process;

use vybe_bytecode::VM;

/// Register the always-on Node modules (`os`, `path`, `process`,
/// `crypto`). Capability-gated modules (`fs`, `child_process`) live
/// behind the caller's capability check in [`crate::modules`].
pub fn register_always_on(vm: &mut VM) {
    os::register(vm);
    path::register(vm);
    process::register(vm);
    crypto::register(vm);
}
