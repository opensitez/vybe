//! # ecma:fixedarray builtins
//!
//! `FixedArray` is the Vybe name for a thin wrapper around a
//! spec-correct WASM GC fixed-length array (`array.new_fixed` / the
//! underlying `(ref array<externref>)` object) — used by:
//!   - COBOL `OCCURS n TIMES` (fixed-size tables)
//!   - VB `Dim arr(5) as Integer` without `ReDim`
//!   - Python `tuple`
//!   - C# `T[]` when the compiler proves non-growable
//!   - Any interop boundary that wants pure-spec GC arrays
//!
//! This module exists so compilers that need to hand a fixed-length
//! array to code expecting a growable `Array` (and vice versa) don't
//! have to hand-roll byte-level conversions. All imports fall
//! through to the backing `ObjectKind::Array` representation —
//! there's no dedicated `ObjectKind::FixedArray` variant because
//! **fixed-vs-growable is a compile-time intent, not a runtime
//! type**. The compiler emits `ecma:fixedarray.freeze` to
//! signal "treat this as immutable length from here on" and we
//! honor that at the mutation sites.
//!
//! See `JS_BUILTIN_CONVENTIONS.md`.

use crate::encoding::*;

pub const MODULE: &str = "ecma:fixedarray";

pub const IMPORTS: &[&str] = &[
    "newWithLength", // FixedArray(n) — n null elements
    "fromArray",     // FixedArray from growable Array (snapshot)
    "toArray",       // Convert FixedArray → growable Array
    "length",        // Read length (same shape as Array.length)
    "get",           // Element access by index
    "isFixedArray",  // Tagged check (distinguishes from growable Array)
    "freeze",        // Tag an Array as fixed — mutations trap
    "isFrozen",      // Check whether frozen
];

pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "newWithLength" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "fromArray" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "toArray" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "length" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "get" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "isFixedArray" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "freeze" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "isFrozen" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        _ => return false,
    }
    true
}
