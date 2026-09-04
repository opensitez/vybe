//! `wasm:js-*` host imports — real WebAssembly CG proposals.
//!
//! - [`crate::js_string_builtins`]    — merged js-string-builtins (V8 native),
//!   incl. the `wasm:js-string` extensions from the Stage-1
//!   js-primitive-builtins proposal.
//! - [`crate::js_primitive_builtins`] — Stage-1 js-primitive-builtins
//!   (`wasm:js-{number,boolean,undefined,symbol,bigint}`).
//!
//! These live alongside the other WASM proposal implementations in
//! `vybe_runtime` (`simd`, `jspi`, `component`, …).

use crate::VM;

/// Register the `wasm:js-string` + `wasm:js-{number,boolean,…}` builtin
/// host functions on the VM.
pub fn register(vm: &mut VM) {
    crate::js_string_builtins::register(vm);
    crate::js_primitive_builtins::register(vm);
    crate::js_prototype_builtins::register(vm);
}
