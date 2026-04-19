//! # gc proposal
//!
//! Spec: <https://github.com/WebAssembly/gc>. Adds garbage-collected
//! reference types (`struct`, `array`, `i31ref`, `anyref`, …) + casts
//! + struct/array ops. Prefix `0xFB`.
//!
//! ## Status in Vybe
//!
//! | Feature                | Status | Notes |
//! |------------------------|--------|-------|
//! | `struct.new`           | ✅ | `STRUCT_NEW` opcode |
//! | `struct.new_default`   | ✅ | `STRUCT_NEW_DEFAULT` |
//! | `struct.get` / `set`   | ✅ | field-indexed; VM uses property bag fallback for non-typed |
//! | `array.new_fixed`      | ✅ | `ARRAY_NEW_FIXED` |
//! | `array.get` / `set`    | ✅ | bounds-checked |
//! | `array.len`            | ✅ | `ARRAY_LEN` |
//! | `ref.test`             | ✅ | uses per-type id table |
//! | `ref.cast`             | ✅ | traps on failure |
//! | `br_on_cast` / `br_on_cast_fail` | ✅ | combined with block depth |
//! | `anyref` / `any.convert_extern` / `extern.convert_any` | ✅ | externref↔anyref via our universal ext carrier |
//! | `i31.new` / `i31.get_s` / `i31.get_u` | ✅ | boxed integer fast path |
//! | `ref.eq` (`0xFB 0x13`) | ✅ | `REF_EQ` opcode — identity check for Objects / Symbols / Strings / null |
//! | Shared GC objects      | ⚠  | `SHARED_NEW` / `SHARED_STRUCT_*` opcodes exist; limited VM support |
//!
//! ## Emitter
//!
//! The 0xFB-prefix emitter lives in `code.rs::emit_gc_op` (access to
//! type context needed for struct/array type indices).

use crate::Chunk;

/// GC proposal declares no imports — types live in the type section.
pub const IMPORTS: &[(&str, &str)] = &[];
/// No globals either.
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];

pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> { Vec::new() }
