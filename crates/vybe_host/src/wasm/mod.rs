//! # `wasm:*` host imports — real WebAssembly CG proposals.
//!
//! Each file corresponds to one proposal.
//!
//! - `js_string_builtins`   — **merged** js-string-builtins (V8 native);
//!                            also includes the wasm:js-string extensions
//!                            from the Stage-1 js-primitive-builtins proposal
//! - `js_primitive_builtins`— Stage-1 js-primitive-builtins
//!                            (wasm:js-{number,boolean,undefined,symbol,bigint})
//!
//! See `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md` for the
//! marshaling + error-handling contract every entry must satisfy.

pub mod js_string_builtins;
pub mod js_primitive_builtins;

use vybe_bytecode::VM;

pub fn register(vm: &mut VM) {
    js_string_builtins::register(vm);
    js_primitive_builtins::register(vm);
}
