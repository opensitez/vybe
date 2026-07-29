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

pub mod build;
pub mod c_runtime;
pub mod dispatch;
pub mod math_adapter;
pub mod math_runtime;
pub mod posix_adapter;
pub mod regex_adapter;
pub mod stdio_adapter;
pub mod stdio_format;
pub mod stdlib_adapter;
pub mod stdlib_runtime;
pub mod string_adapter;
pub mod string_runtime;
pub mod thread_adapter;
pub mod time_adapter;
pub mod tree_register;
pub mod uchar_adapter;
pub mod wchar_adapter;
