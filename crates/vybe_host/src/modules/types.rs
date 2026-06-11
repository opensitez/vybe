//! Built-in .NET collection types — retired.
//!
//! All vybe:types host fns replaced by compile-time adapters:
//!   List/Queue/Stack   → `collections.new` (ARRAY_NEW) + `collections.*` range ops
//!   Dictionary         → `ecma:map.*`
//!   HashSet            → `ecma:set.*`
//!   Range ops          → `emitter/collections.rs` helpers
//!
//! TypeDef constructors in `builtin_types.rs` point at `ecma:array.new`,
//! `ecma:map.new`, and `ecma:set.new` directly.

use vybe_bytecode::VM;

pub fn register(_vm: &mut VM) {}
