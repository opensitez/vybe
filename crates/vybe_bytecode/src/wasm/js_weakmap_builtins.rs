//! # ecma:weakmap and ecma:weakset builtins
//!
//! Host imports for `WeakMap.prototype.*` / `WeakSet.prototype.*`
//! per ECMA-262 §24.3 / §24.4.
//!
//! WASM GC MVP does not yet have true weak references. Vybe VM uses
//! the `weak_table` crate for real weak semantics (Phase B3 handlers);
//! the pure-WASM polyfill falls back to strong references (leaks,
//! but functionally correct) until WASM GC Post-MVP lands.

use super::encoding::*;

pub const WEAKMAP_MODULE: &str = "ecma:weakmap";
pub const WEAKSET_MODULE: &str = "ecma:weakset";

pub const WEAKMAP_IMPORTS: &[&str] = &["new", "fromIterable", "get", "set", "has", "delete"];

pub const WEAKSET_IMPORTS: &[&str] = &["new", "fromIterable", "add", "has", "delete"];

pub fn write_weakmap_signature(out: &mut Vec<u8>, name: &str) -> bool {
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
        "get" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "set" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
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
        _ => return false,
    }
    true
}

pub fn write_weakset_signature(out: &mut Vec<u8>, name: &str) -> bool {
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
        _ => return false,
    }
    true
}
