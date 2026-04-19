//! # js-primitive-builtins proposal
//!
//! Spec: <https://github.com/WebAssembly/js-primitive-builtins>. Extends
//! js-string-builtins with imports for the remaining JS primitives:
//! `number`, `boolean`, `undefined`, `symbol`, `bigint`. This module is
//! the **single source of truth** for everything under the modules:
//!
//! * `wasm:js-number`      — numeric coercions + type tests
//! * `wasm:js-boolean`     — `true` / `false` globals + cast / test
//! * `wasm:js-undefined`   — `void 0` global + type test
//! * `wasm:js-symbol`      — identity test + equals
//! * `wasm:js-bigint`      — type test
//!
//! ## Functions
//!
//! | Module         | Import                                  |
//! |----------------|-----------------------------------------|
//! | js-number      | fromF64, fromI32, fromU32               |
//! | js-number      | toF64, toI32, toU32                     |
//! | js-number      | test, testI32, testU32                  |
//! | js-boolean     | test, cast                              |
//! | js-undefined   | test                                    |
//! | js-symbol      | test, equals                            |
//! | js-bigint      | test                                    |
//!
//! ## Globals
//!
//! The proposal recommends importing singleton values as WASM globals
//! rather than calling a constructor every time:
//!
//! | Module         | Global name | Indexed as               |
//! |----------------|-------------|--------------------------|
//! | js-undefined   | `value`     | `JS_GLOBAL_UNDEFINED` (0) |
//! | js-boolean     | `true`      | `JS_GLOBAL_TRUE`      (1) |
//! | js-boolean     | `false`     | `JS_GLOBAL_FALSE`     (2) |

use super::encoding::*;

/// Module + function imports, ordered so the emitted import section is
/// stable across builds.
pub const FUNC_IMPORTS: &[(&str, &str)] = &[
    // js-number — full stage-1 surface
    ("wasm:js-number", "fromF64"),
    ("wasm:js-number", "fromI32"),
    ("wasm:js-number", "fromU32"),
    ("wasm:js-number", "toF64"),
    ("wasm:js-number", "toI32"),
    ("wasm:js-number", "toU32"),
    ("wasm:js-number", "test"),
    ("wasm:js-number", "testI32"),
    ("wasm:js-number", "testU32"),
    // js-boolean
    ("wasm:js-boolean", "test"),
    ("wasm:js-boolean", "cast"),
    // js-undefined
    ("wasm:js-undefined", "test"),
    // js-symbol
    ("wasm:js-symbol", "test"),
    ("wasm:js-symbol", "equals"),
    // js-bigint
    ("wasm:js-bigint", "test"),
];

/// Global imports — indices here MUST match the `JS_GLOBAL_*` constants
/// exposed from `sections.rs` so the emitter can reference them.
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[
    ("wasm:js-undefined", "value"),
    ("wasm:js-boolean",   "true"),
    ("wasm:js-boolean",   "false"),
];

/// Emit the WASM function signature for a `(module, name)` pair.
/// Returns `true` when the pair is recognised. The caller has already
/// pushed the `TYPE_FUNC` tag byte.
pub fn write_signature(out: &mut Vec<u8>, module: &str, name: &str) -> bool {
    match (module, name) {
        // ── js-number ──────────────────────────────────────────────
        ("wasm:js-number", "fromI32") | ("wasm:js-number", "fromU32") => {
            write_leb128_u32(out, 1); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        ("wasm:js-number", "fromF64") => {
            write_leb128_u32(out, 1); out.push(TYPE_F64);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        ("wasm:js-number", "toI32") | ("wasm:js-number", "toU32") => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        ("wasm:js-number", "toF64") => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_F64);
        }
        ("wasm:js-number", "test")
        | ("wasm:js-number", "testI32")
        | ("wasm:js-number", "testU32")
        // ── js-boolean ─────────────────────────────────────────────
        | ("wasm:js-boolean", "test")
        | ("wasm:js-boolean", "cast")
        // ── js-undefined ───────────────────────────────────────────
        | ("wasm:js-undefined", "test")
        // ── js-symbol ──────────────────────────────────────────────
        | ("wasm:js-symbol", "test")
        // ── js-bigint ──────────────────────────────────────────────
        | ("wasm:js-bigint", "test") => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        ("wasm:js-symbol", "equals") => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        _ => return false,
    }
    true
}
