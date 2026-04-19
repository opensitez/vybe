//! # wasm:js-typedarray-* builtins
//!
//! Host imports for the 11 typed-array variants per ECMA-262 §23.2.
//!
//! Each variant gets its own module name:
//!   - `wasm:js-int8array`      (`Int8Array`)
//!   - `wasm:js-uint8array`     (`Uint8Array`)
//!   - `wasm:js-uint8clamped`   (`Uint8ClampedArray`)
//!   - `wasm:js-int16array`     (`Int16Array`)
//!   - `wasm:js-uint16array`    (`Uint16Array`)
//!   - `wasm:js-int32array`     (`Int32Array`)
//!   - `wasm:js-uint32array`    (`Uint32Array`)
//!   - `wasm:js-float32array`   (`Float32Array`)
//!   - `wasm:js-float64array`   (`Float64Array`)
//!   - `wasm:js-bigint64array`  (`BigInt64Array`)
//!   - `wasm:js-biguint64array` (`BigUint64Array`)
//!
//! All variants share the same method surface; only the element type
//! differs. Signatures below are generic on the element type, with
//! per-variant specialization via the `TypedElem` enum.

use super::encoding::*;

/// Element type of a typed-array variant — determines the WASM type
/// used for get/set value arguments.
#[derive(Copy, Clone, Debug)]
pub enum TypedElem {
    /// Int8Array — i32 with sign-extension on get
    I8,
    /// Uint8Array, Uint8ClampedArray — i32 with zero-extension on get
    U8,
    /// Uint8ClampedArray — i32 with saturating-clamp on set
    U8Clamped,
    /// Int16Array — i32 sign-ext
    I16,
    /// Uint16Array — i32 zero-ext
    U16,
    /// Int32Array — i32
    I32,
    /// Uint32Array — i32 (unsigned interpretation)
    U32,
    /// Float32Array — f32
    F32,
    /// Float64Array — f64
    F64,
    /// BigInt64Array — i64
    BigI64,
    /// BigUint64Array — i64 (unsigned)
    BigU64,
}

impl TypedElem {
    /// Returns (module_name, bytes_per_element, value_wasm_type).
    pub fn info(self) -> (&'static str, u32, u8) {
        match self {
            TypedElem::I8       => ("wasm:js-int8array",     1, TYPE_I32),
            TypedElem::U8       => ("wasm:js-uint8array",    1, TYPE_I32),
            TypedElem::U8Clamped=> ("wasm:js-uint8clamped",  1, TYPE_I32),
            TypedElem::I16      => ("wasm:js-int16array",    2, TYPE_I32),
            TypedElem::U16      => ("wasm:js-uint16array",   2, TYPE_I32),
            TypedElem::I32      => ("wasm:js-int32array",    4, TYPE_I32),
            TypedElem::U32      => ("wasm:js-uint32array",   4, TYPE_I32),
            TypedElem::F32      => ("wasm:js-float32array",  4, TYPE_F32),
            TypedElem::F64      => ("wasm:js-float64array",  8, TYPE_F64),
            TypedElem::BigI64   => ("wasm:js-bigint64array", 8, TYPE_I64),
            TypedElem::BigU64   => ("wasm:js-biguint64array",8, TYPE_I64),
        }
    }

    pub fn module(self) -> &'static str { self.info().0 }
    pub fn bytes_per_element(self) -> u32 { self.info().1 }
    pub fn value_type(self) -> u8 { self.info().2 }
}

/// Ordered list of all 11 typed-array variants.
pub const VARIANTS: &[TypedElem] = &[
    TypedElem::I8,
    TypedElem::U8,
    TypedElem::U8Clamped,
    TypedElem::I16,
    TypedElem::U16,
    TypedElem::I32,
    TypedElem::U32,
    TypedElem::F32,
    TypedElem::F64,
    TypedElem::BigI64,
    TypedElem::BigU64,
];

/// Method surface — every typed-array variant exports these methods.
/// Matches ECMA-262 §23.2.3 (TypedArray prototype).
pub const IMPORTS: &[&str] = &[
    // ── Construction ────────────────────────────────────────────────
    "newWithLength",             // new TypedArray(length)
    "newFromBuffer",             // new TypedArray(buffer, byteOffset?, length?)
    "newFromIterable",           // new TypedArray(iterable)
    "newFromTypedArray",         // new TypedArray(otherTypedArray) — copies

    // ── Statics ─────────────────────────────────────────────────────
    "from",                      // TypedArray.from(source, mapFn?)
    "of",                        // TypedArray.of(...values)

    // ── Properties ──────────────────────────────────────────────────
    "buffer",                    // arr.buffer
    "byteOffset",                // arr.byteOffset
    "byteLength",                // arr.byteLength
    "length",                    // arr.length

    // ── Element access ──────────────────────────────────────────────
    "get",                       // arr[i]
    "set",                       // arr[i] = v
    "at",                        // arr.at(i)

    // ── Mutating ────────────────────────────────────────────────────
    "setArray",                  // arr.set(sourceArray, offset?)
    "copyWithin",
    "fill",
    "reverse",
    "sort",

    // ── Non-mutating ────────────────────────────────────────────────
    "slice",
    "subarray",                  // view over same buffer, no copy
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
    "with",
];

/// Emit the WASM function signature for the given import on the given
/// typed-array variant. Returns `true` when the name is recognised.
pub fn write_signature(out: &mut Vec<u8>, variant: TypedElem, name: &str) -> bool {
    let value_t = variant.value_type();
    match name {
        "newWithLength" => {
            write_leb128_u32(out, 1); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "newFromBuffer" => {
            // (buffer, byteOffset, length) — pass -1 for omitted
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "newFromIterable" | "newFromTypedArray" | "from" | "of" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "buffer" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "byteOffset" | "byteLength" | "length" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "get" | "at" => {
            // (arr, index) -> element_value
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(value_t);
        }
        "set" => {
            // (arr, index, value) -> ()
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(value_t);
            write_leb128_u32(out, 0);
        }
        "setArray" => {
            // (arr, source, offset) -> ()
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        "copyWithin" => {
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "fill" => {
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF); out.push(value_t); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "reverse" | "toReversed" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "sort" | "toSorted" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "slice" | "subarray" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "indexOf" | "lastIndexOf" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(value_t); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "includes" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(value_t); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "find" | "findLast" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(value_t);
        }
        "findIndex" | "findLastIndex" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "join" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "toString" | "toLocaleString" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "forEach" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 0);
        }
        "map" | "filter" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "reduce" | "reduceRight" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "some" | "every" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "keys" | "values" | "entries" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "with" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(value_t);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        _ => return false,
    }
    true
}
