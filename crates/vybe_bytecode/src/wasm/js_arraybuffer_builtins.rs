//! # wasm:js-arraybuffer, wasm:js-sharedarraybuffer, wasm:js-dataview
//!
//! Host imports for:
//!   - `ArrayBuffer.*` / `ArrayBuffer.prototype.*` per ECMA-262 §25.1
//!   - `SharedArrayBuffer.*` per §25.2
//!   - `DataView.*` per §25.3
//!
//! All three live in this file because they're intimately related —
//! `DataView` is a view over an `ArrayBuffer`, and
//! `SharedArrayBuffer` is a near-twin of `ArrayBuffer` with
//! thread-safety.

use super::encoding::*;

// ── ArrayBuffer ──────────────────────────────────────────────────────

pub const ARRAYBUFFER_MODULE: &str = "wasm:js-arraybuffer";

pub const ARRAYBUFFER_IMPORTS: &[&str] = &[
    "new",                       // new ArrayBuffer(byteLength)
    "newResizable",              // new ArrayBuffer(byteLength, { maxByteLength })
    "byteLength",                // ab.byteLength
    "maxByteLength",             // ab.maxByteLength
    "resizable",                 // ab.resizable
    "detached",                  // ab.detached
    "slice",                     // ab.slice(start, end)
    "resize",                    // ab.resize(newByteLength)
    "transfer",                  // ab.transfer(newByteLength?)
    "transferToFixedLength",     // ab.transferToFixedLength(newByteLength?)
    "isView",                    // ArrayBuffer.isView(v) — static
];

pub fn write_arraybuffer_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "new" => {
            write_leb128_u32(out, 1); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "newResizable" => {
            // (byteLength, maxByteLength) -> ArrayBuffer
            write_leb128_u32(out, 2);
            out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "byteLength" | "maxByteLength" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "resizable" | "detached" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "slice" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "resize" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        "transfer" | "transferToFixedLength" => {
            // (ab, newByteLength) -> new_ArrayBuffer
            // Passing -1 means "keep current length" (saves declaring
            // an optional-arg variant).
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "isView" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        _ => return false,
    }
    true
}

// ── SharedArrayBuffer ────────────────────────────────────────────────

pub const SHAREDARRAYBUFFER_MODULE: &str = "wasm:js-sharedarraybuffer";

pub const SHAREDARRAYBUFFER_IMPORTS: &[&str] = &[
    "new",                       // new SharedArrayBuffer(byteLength)
    "newGrowable",               // new SharedArrayBuffer(byteLength, { maxByteLength })
    "byteLength",
    "maxByteLength",
    "growable",
    "slice",
    "grow",                      // sab.grow(newByteLength)
];

pub fn write_sharedarraybuffer_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "new" => {
            write_leb128_u32(out, 1); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "newGrowable" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "byteLength" | "maxByteLength" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "growable" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        "slice" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "grow" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        _ => return false,
    }
    true
}

// ── DataView ─────────────────────────────────────────────────────────

pub const DATAVIEW_MODULE: &str = "wasm:js-dataview";

pub const DATAVIEW_IMPORTS: &[&str] = &[
    "new",                       // new DataView(buffer, byteOffset?, byteLength?)
    "buffer",                    // dv.buffer
    "byteOffset",                // dv.byteOffset
    "byteLength",                // dv.byteLength
    // Getters
    "getInt8",
    "getUint8",
    "getInt16",
    "getUint16",
    "getInt32",
    "getUint32",
    "getBigInt64",
    "getBigUint64",
    "getFloat32",
    "getFloat64",
    // Setters
    "setInt8",
    "setUint8",
    "setInt16",
    "setUint16",
    "setInt32",
    "setUint32",
    "setBigInt64",
    "setBigUint64",
    "setFloat32",
    "setFloat64",
];

pub fn write_dataview_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "new" => {
            // (buffer, byteOffset, byteLength) -> DataView
            // byteOffset/byteLength -1 means "omitted" (default).
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "buffer" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
        }
        "byteOffset" | "byteLength" => {
            write_leb128_u32(out, 1); out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        // 8-bit getters — no endianness arg
        "getInt8" | "getUint8" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        // Multi-byte int getters — (view, offset, littleEndian) -> i32
        "getInt16" | "getUint16" | "getInt32" | "getUint32" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I32);
        }
        // 64-bit int getters
        "getBigInt64" | "getBigUint64" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_I64);
        }
        // Float getters
        "getFloat32" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_F32);
        }
        "getFloat64" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 1); out.push(TYPE_F64);
        }
        // 8-bit setters — no endianness arg
        "setInt8" | "setUint8" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        // Multi-byte int setters — (view, offset, value, littleEndian)
        "setInt16" | "setUint16" | "setInt32" | "setUint32" => {
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I32); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        // 64-bit int setters
        "setBigInt64" | "setBigUint64" => {
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_I64); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        // Float setters
        "setFloat32" => {
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_F32); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        "setFloat64" => {
            write_leb128_u32(out, 4);
            out.push(TYPE_EXTERNREF); out.push(TYPE_I32); out.push(TYPE_F64); out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        _ => return false,
    }
    true
}
