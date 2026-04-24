//! # vybe:js-structured-clone builtin
//!
//! Host import that satisfies the HTML `structuredClone(value)`
//! algorithm — deep copy across Array / Object / Map / Set /
//! ArrayBuffer / TypedArray / DataView / primitives.
//!
//! Needed by: JS `structuredClone`, Python `copy.deepcopy`, Worker
//! `postMessage` (the serialization path), Ruby `Marshal.load(
//! Marshal.dump(x))` equivalent.

use super::encoding::*;

pub const MODULE: &str = "vybe:js-structured-clone";

pub const IMPORTS: &[&str] = &[
    "clone",   // structuredClone(value, options?)
];

pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "clone" => {
            // (value) -> clone  (options ignored — MVP doesn't
            // implement transfer lists)
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        _ => return false,
    }
    true
}
