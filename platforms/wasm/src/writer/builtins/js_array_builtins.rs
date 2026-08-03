//! # ecma:array builtins
//!
//! Host imports that satisfy `Array.prototype.*` and `Array.*` static
//! methods per ECMA-262 §23.1. Vybe VM registers native Rust handlers
//! (see `vybe_host::modules::array` — Phase B3); browsers satisfy the
//! imports via JS glue (see `tools/vybe-loader/vybe_js_collections.js`
//! — Phase C); plain-WASM engines link the polyfill module that ships
//! with `vybe build --target portable` (Phase C).
//!
//! Marshaling conventions pinned in `JS_BUILTIN_CONVENTIONS.md`:
//!   - Array instances → `externref`
//!   - Indices / sizes / counts → `i32`
//!   - Booleans → `i32` (0 = false, non-zero = true)
//!   - Element values → `externref` (universal value carrier)
//!   - Callback funcs → `externref` (dispatched by the host;
//!     consistent with our dynamic value ABI)
//!
//! Spec reference: <https://tc39.es/ecma262/#sec-array-objects>.

use crate::encoding::*;

pub const MODULE: &str = "ecma:array";

/// Every `ecma:array` import. Matches ECMA-262 §23.1 surface.
pub const IMPORTS: &[&str] = &[
    // ── Constructors / statics ──────────────────────────────────────
    "new",           // new Array()
    "newWithLength", // new Array(n)
    "of",            // Array.of(...values) — runtime-variadic; handler walks stack
    "from",          // Array.from(iterable, mapFn?)
    "fromAsync",     // Array.fromAsync(asyncIterable, mapFn?) — ES2024
    "isArray",       // Array.isArray(v)
    // ── Property access ─────────────────────────────────────────────
    "get",       // arr[i] (spec: OrdinaryGet with integer key)
    "set",       // arr[i] = v
    "length",    // arr.length (getter)
    "setLength", // arr.length = n (truncate or null-fill extend)
    "at",        // arr.at(i) — negative indices handled at language layer
    // ── Mutating prototype methods ──────────────────────────────────
    "push",
    "pop",
    "shift",
    "unshift",
    "splice",
    "reverse",
    "sort",
    "fill",
    "copyWithin",
    // ── Non-mutating prototype methods ──────────────────────────────
    "slice",
    "concat",
    "indexOf",
    "lastIndexOf",
    "includes",
    "find",
    "findIndex",
    "findLast",
    "findLastIndex",
    "join",
    "toString",
    "toLocaleString",
    "flat",
    "flatMap",
    // ── Iteration / higher-order ────────────────────────────────────
    "forEach",
    "map",
    "filter",
    "reduce",
    "reduceRight",
    "some",
    "every",
    "keys",
    "values",
    "entries",
    // ── ES2023 non-mutating variants ────────────────────────────────
    "toReversed",
    "toSorted",
    "toSpliced",
    "with",
    // ── ES2025 group-by ─────────────────────────────────────────────
    "group",
    "groupToMap",
];

/// Emit the WASM function signature for the given import. Returns
/// `true` when the name is recognised. The caller has already pushed
/// the `TYPE_FUNC` tag byte.
pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        // ── Constructors ────────────────────────────────────────────
        "new" => {
            write_leb128_u32(out, 0);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "newWithLength" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "of" | "from" | "fromAsync" => {
            // `of(...values)` / `from(iterable, mapFn?)` — the handler
            // reads variadic args off the WASM stack using a
            // caller-supplied count + array of externrefs.
            // Signature: (count: i32, args_array: externref) -> externref
            write_leb128_u32(out, 2);
            out.push(TYPE_I32);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "isArray" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }

        // ── Property access ─────────────────────────────────────────
        "get" => {
            // (arr, index) -> value
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "set" => {
            // (arr, index, value) -> ()
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        "length" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "setLength" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        "at" => {
            // (arr, index) -> value (undefined if OOB — caller layer
            // translates negative indices before calling)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }

        // ── Mutating prototype methods ──────────────────────────────
        "push" => {
            // (arr, value) -> new_length
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "pop" | "shift" => {
            // (arr) -> popped_value (undefined if empty)
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "unshift" => {
            // (arr, value) -> new_length
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "splice" => {
            // (arr, start, deleteCount, items_array) -> deleted_array
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "reverse" | "toReversed" => {
            // reverse: in-place, returns self; toReversed: new array
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "sort" | "toSorted" => {
            // (arr, compareFn: externref or null) -> sorted_arr
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "fill" => {
            // (arr, value, start, end) -> self
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "copyWithin" => {
            // (arr, target, start, end) -> self
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }

        // ── Non-mutating slicing ────────────────────────────────────
        "slice" => {
            // (arr, start, end) -> new_arr
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "concat" => {
            // (arr, other) -> new_arr
            // Spec takes variadic; compiler lowers to pairwise concat
            // when multiple args are present.
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "indexOf" | "lastIndexOf" => {
            // (arr, value, fromIndex) -> i32 (-1 if not found)
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "includes" => {
            // (arr, value, fromIndex) -> bool (i32)
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "find" | "findLast" => {
            // (arr, predicate) -> found_value or undefined
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "findIndex" | "findLastIndex" => {
            // (arr, predicate) -> i32 (-1 if not found)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "join" => {
            // (arr, separator) -> string (externref)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "toString" | "toLocaleString" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "flat" => {
            // (arr, depth) -> new_arr
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "flatMap" => {
            // (arr, mapFn) -> new_arr
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }

        // ── Iteration / higher-order ────────────────────────────────
        "forEach" => {
            // (arr, callback) -> ()
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        "map" | "filter" => {
            // (arr, callback) -> new_arr
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "reduce" | "reduceRight" => {
            // (arr, callback, initialValue) -> accumulated_value
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "some" | "every" => {
            // (arr, predicate) -> bool (i32)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "keys" | "values" | "entries" => {
            // (arr) -> iterator (externref wrapping an Iter object
            // per the JS @@iterator protocol)
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }

        // ── ES2023 non-mutating variants ────────────────────────────
        "toSpliced" => {
            // (arr, start, deleteCount, items_array) -> new_arr
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "with" => {
            // (arr, index, value) -> new_arr
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_I32);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }

        // ── ES2025 group-by ────────────────────────────────────────
        "group" | "groupToMap" => {
            // (arr, callback) -> Object or Map grouping by key
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }

        _ => return false }
    true
}
