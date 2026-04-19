//! # js-string-builtins proposal
//!
//! Spec: <https://github.com/WebAssembly/js-string-builtins>. Standard
//! imports under the `wasm:js-string` namespace that let WASM code
//! manipulate JS strings without round-tripping through glue code.
//!
//! This module is the **single source of truth** for:
//! * the list of `wasm:js-string` imports we declare
//! * their (param → result) signatures
//!
//! Anywhere else that needs this information queries this module rather
//! than maintaining a parallel list.
//!
//! ## Spec coverage
//!
//! | Op                   | Imported | Emitter wires it? |
//! |----------------------|----------|-------------------|
//! | `test`               | ✅        | ✅ `REF_IS_STRING`, `REF_TYPEOF` |
//! | `cast`               | ✅        | ☐ (declared; no opcode lowers to it yet) |
//! | `concat`             | ✅        | ✅ `STR_CONCAT` |
//! | `equals`             | ✅        | ✅ `STR_EQUALS` |
//! | `compare`            | ✅        | ✅ `STR_COMPARE` |
//! | `length`             | ✅        | ✅ `STR_LENGTH` |
//! | `charCodeAt`         | ✅        | ✅ `STR_CHAR_CODE_AT` |
//! | `codePointAt`        | ✅        | ✅ `STR_CODE_POINT_AT` |
//! | `fromCharCode`       | ✅        | ✅ `STR_FROM_CHAR_CODE` |
//! | `fromCodePoint`      | ✅        | ✅ `STR_FROM_CODE_POINT` |
//! | `substring`          | ✅        | ✅ `STR_SUBSTRING`, `STR_SLICE` |
//! | `intoCharCodeArray`  | ✅        | ✅ `STR_INTO_CHAR_CODES` |
//! | `fromCharCodeArray`  | ✅        | ✅ `STR_FROM_CHAR_CODES` |
//!
//! Plus, from the js-primitive-builtins extension:
//! | `fromI32`, `fromU32`, `fromI64`, `fromU64`, `fromF64` | ✅ | ☐ (declared; no opcode lowers yet) |

use super::encoding::*;

pub const MODULE: &str = "wasm:js-string";

/// All `wasm:js-string` imports declared by the emitter.
pub const IMPORTS: &[&str] = &[
    "test", "cast",
    "concat", "equals", "compare", "length",
    "charCodeAt", "codePointAt",
    "fromCharCode", "fromCodePoint",
    "substring",
    "intoCharCodeArray", "fromCharCodeArray",
    // js-primitive-builtins: numeric-to-string formatting
    "fromI32", "fromU32", "fromI64", "fromU64", "fromF64",
];

/// Emit the WASM function signature for the given import, appending to
/// `out`. Returns `true` when the name is recognised. The caller has
/// already pushed the `TYPE_FUNC` tag byte.
pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "test" | "length" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "cast" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "concat" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "equals" | "compare" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "charCodeAt" | "codePointAt" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "fromCharCode" | "fromCodePoint" | "fromI32" | "fromU32" => {
            write_leb128_u32(out, 1); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "fromI64" | "fromU64" => {
            write_leb128_u32(out, 1); out.push(TYPE_I64);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "fromF64" => {
            write_leb128_u32(out, 1); out.push(TYPE_F64);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "substring" | "fromCharCodeArray" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "intoCharCodeArray" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        _ => return false,
    }
    true
}
