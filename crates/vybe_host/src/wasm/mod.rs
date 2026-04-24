//! # `wasm:*` host imports — real WebAssembly CG proposals.
//!
//! This folder hosts Rust implementations of imports defined by actual
//! WebAssembly proposals (merged or staged through the CG process). Only
//! names that appear in a real Overview.md under `proposals/` may live
//! here — everything else that looks "WASM-ish" is Vybe-invented and
//! belongs under `ecma/` or another honest namespace.
//!
//! Current entries:
//! - `wasm:js-string`       (**merged** js-string-builtins; V8 native)
//! - `wasm:js-number`       (stage-1 js-primitive-builtins)
//! - `wasm:js-boolean`      (stage-1 js-primitive-builtins)
//! - `wasm:js-undefined`    (stage-1 js-primitive-builtins)
//!
//! See `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md` for
//! the marshaling + error-handling contract every entry must satisfy.

pub mod js_string;
pub mod js_number;
pub mod js_boolean;
pub mod js_undefined;

use vybe_bytecode::VM;

/// Register every `wasm:*` host fn on the VM.
pub fn register(vm: &mut VM) {
    js_string::register(vm);
    js_number::register(vm);
    js_boolean::register(vm);
    js_undefined::register(vm);
}
