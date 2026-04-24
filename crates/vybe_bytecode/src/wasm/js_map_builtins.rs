//! # vybe:js-map builtins
//!
//! Host imports for `Map.prototype.*` and `Map.*` per ECMA-262 §24.1.
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use super::encoding::*;

pub const MODULE: &str = "vybe:js-map";

pub const IMPORTS: &[&str] = &[
    "new",       // new Map()
    "fromEntries", // new Map(iterable)
    "get",       // map.get(key)
    "set",       // map.set(key, value)
    "has",       // map.has(key)
    "delete",    // map.delete(key)
    "clear",     // map.clear()
    "size",      // map.size (getter)
    "keys",      // map.keys()
    "values",    // map.values()
    "entries",   // map.entries()
    "forEach",   // map.forEach(callback)
    "groupBy",   // Map.groupBy(iterable, fn) — ES2025
];

pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "new" => {
            write_leb128_u32(out, 0);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "fromEntries" | "groupBy" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "get" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "set" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "has" | "delete" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "clear" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        "size" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "keys" | "values" | "entries" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "forEach" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        _ => return false,
    }
    true
}
