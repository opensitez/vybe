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

pub mod assert;
pub mod async_hooks;
pub mod buffer;
pub mod child_process;
pub mod crypto;
pub mod dgram;
pub mod dns;
pub mod events;
pub mod fs;
pub mod http;
pub mod https;
pub mod net;
pub mod os;
pub mod path;
pub mod perf_hooks;
pub mod process;
pub mod querystring;
pub mod readline;
pub mod stream;
pub mod string_decoder;
pub mod timers;
pub mod tty;
pub mod url;
pub mod util;
pub mod vm_module;
pub mod worker_threads;
pub mod zlib;

use vybe_runtime::VM;

/// Register the always-on Node modules. Capability-gated modules (`fs`,
/// `child_process`) live behind the caller's capability check in [`crate::modules`].
pub fn register_always_on(vm: &mut VM) {
    assert::register(vm);
    async_hooks::register(vm);
    buffer::register(vm);
    dgram::register(vm);
    dns::register(vm);
    events::register(vm);
    os::register(vm);
    path::register(vm);
    perf_hooks::register(vm);
    process::register(vm);
    crypto::register(vm);
    net::register(vm);
    querystring::register(vm);
    readline::register(vm);
    stream::register(vm);
    string_decoder::register(vm);
    timers::register(vm);
    tty::register(vm);
    url::register(vm);
    util::register(vm);
    vm_module::register(vm);
    worker_threads::register(vm);
    zlib::register(vm);
}

pub mod plugin;
pub use plugin::Plugin;
