//! libc platform adapter — shared by C, Go, and any language targeting
//! libc-compatible WASM (Rust via libc crate, Fortran via libgfortran, etc.).
//!
//! This is the compiler-side adapter layer that maps libc semantics onto the
//! real runtime capabilities available in the Vybe WASM VM: WASM opcodes,
//! `ecma:*`, `wasi:*`, and `vybe:*` host functions.
//!
//! All modules expose pure AST-constructor functions (return `Expression` or
//! `Statement` nodes) so walkers stay language-specific while runtime
//! semantics are centralised here.

pub mod arrays;
pub mod ctype_adapter;
pub mod math_adapter;
pub mod pointers;
pub mod stdio_adapter;
pub mod stdlib_adapter;
pub mod string_adapter;
