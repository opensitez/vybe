//! # ecma:json builtins
//!
//! Host imports that satisfy `JSON.stringify` / `JSON.parse` per
//! ECMA-262 §25.5. Follows the same `wasm:js-*` builtin pattern as
//! the collection modules; see `JS_BUILTIN_CONVENTIONS.md` for
//! marshaling rules.

use crate::encoding::*;

pub const MODULE: &str = "ecma:json";

pub const IMPORTS: &[&str] = &[
    "stringify", // JSON.stringify(value, replacer?, space?)
    "parse",     // JSON.parse(text, reviver?)
];

pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "stringify" => {
            // (value, replacer: externref_or_null, space: externref_or_null)
            //   → externref_string
            // MVP ignores replacer and space; signatures present so
            // callers don't need to pass fewer args than MDN docs.
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "parse" => {
            // (text: externref_string, reviver: externref_or_null)
            //   → externref_value
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        _ => return false,
    }
    true
}
