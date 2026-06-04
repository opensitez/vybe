//! # ecma:set builtins
//!
//! Host imports for `Set.prototype.*` and `Set.*` per ECMA-262 §24.2
//! plus ES2025 set-algebra methods.
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use super::encoding::*;

pub const MODULE: &str = "ecma:set";

pub const IMPORTS: &[&str] = &[
    "new",          // new Set()
    "fromIterable", // new Set(iterable)
    "add",          // set.add(v)
    "has",          // set.has(v)
    "delete",       // set.delete(v)
    "clear",        // set.clear()
    "size",         // set.size
    "values",       // set.values()
    "keys",         // set.keys() (alias for values)
    "entries",      // set.entries()
    "forEach",      // set.forEach(callback)
    // ES2025 set algebra
    "union",
    "intersection",
    "difference",
    "symmetricDifference",
    "isSubsetOf",
    "isSupersetOf",
    "isDisjointFrom",
];

pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "new" => {
            write_leb128_u32(out, 0);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "fromIterable" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "add" => {
            // (set, v) -> set (spec returns the set for chaining)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "has" | "delete" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "clear" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        "size" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "values" | "keys" | "entries" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "forEach" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        "union" | "intersection" | "difference" | "symmetricDifference" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "isSubsetOf" | "isSupersetOf" | "isDisjointFrom" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        _ => return false,
    }
    true
}
