//! WASM binary encoding/decoding primitives.
//! Constants, LEB128, section writing, value serialization.

use std::sync::Arc;
use vybe_runtime::value::Value;

// ── WASM binary format constants ────────────────────────────────────────

pub const WASM_MAGIC: [u8; 4] = [0x00, 0x61, 0x73, 0x6D];
pub const WASM_VERSION: [u8; 4] = [0x01, 0x00, 0x00, 0x00];

// Section IDs
pub const SECTION_CUSTOM: u8 = 0;
pub const SECTION_TYPE: u8 = 1;
pub const SECTION_IMPORT: u8 = 2;
pub const SECTION_FUNCTION: u8 = 3;
pub const SECTION_MEMORY: u8 = 5;
pub const SECTION_GLOBAL: u8 = 6;
pub const SECTION_EXPORT: u8 = 7;
pub const SECTION_CODE: u8 = 10;
pub const SECTION_TAG: u8 = 13; // exception-handling proposal

// Value types
pub const TYPE_FUNC: u8 = 0x60;
pub const TYPE_I32: u8 = 0x7F;
pub const TYPE_I64: u8 = 0x7E;
pub const TYPE_F64: u8 = 0x7C;
pub const TYPE_F32: u8 = 0x7D;
pub const TYPE_EXTERNREF: u8 = 0x6F;
pub const TYPE_FUNCREF: u8 = 0x70;
pub const TYPE_VOID: u8 = 0x40;

// GC type encoding
pub const GC_STRUCT: u8 = 0x5F; // -0x21: struct composite type
pub const GC_ARRAY: u8 = 0x5E; // -0x22: array composite type
pub const GC_SUB: u8 = 0x50; // -0x30: sub (open — further subtyping allowed)
pub const GC_SUB_FINAL: u8 = 0x4F; // -0x31: sub final
pub const GC_REC: u8 = 0x4E; // -0x32: recursive type group
// Custom Descriptors proposal — prefixes that may sit between a subtype's
// supertype vector and its composite type.
pub const CD_DESCRIBES: u8 = 0x4C; // (describes $x)
pub const CD_DESCRIPTOR: u8 = 0x4D; // (descriptor $x)
// `heaptype ::= ... | 0x62 x:u32 => exact x`. The index is a plain `u32`,
// NOT the `s33` every other heaptype uses — deliberately, so that an exact
// ABSTRACT heap type cannot be encoded. As an s33, 0x62 reads back as -30,
// which sits in the reserved negative space, so a leading 0x62 is never
// ambiguous with a typeidx.
pub const HEAPTYPE_EXACT: u8 = 0x62;
// `externtype ::= ... | 0x20 x:typeidx => func exact x`. Bit 6 is reserved
// for marking exactness of other externtype kinds later. Exports never
// declare exactness — an export section using 0x20 is malformed.
pub const EXTERNTYPE_FUNC_EXACT: u8 = 0x20;
pub const GC_MUT: u8 = 0x01; // mutable field
pub const GC_IMMUT: u8 = 0x00; // immutable field

// Packed types — valid only as array/struct field storage type, not
// as a top-level value type. Per WASM GC proposal.
pub const PACKED_I8: u8 = 0x78; // -0x08
pub const PACKED_I16: u8 = 0x77; // -0x09

// Abstract heap types (GC proposal). Used as the heaptype operand of
// `ref.null`, `ref.test`, `ref.cast`, `br_on_cast` — single-byte encodings
// when the target is an abstract type (rather than a concrete typeidx).
// See `proposals/gc/proposals/gc/MVP.md` §Reference types.
//
// Defined in `vybe_runtime::opcode::heaptype`, beside the opcodes that take
// them — `ref.null` cannot be emitted without one — and re-exported here so
// the binary writer and the bytecode emitter cannot disagree about a byte.
pub use vybe_runtime::opcode::heaptype::{
    HT_ANY, HT_ARRAY, HT_EQ, HT_EXTERN, HT_FUNC, HT_I31, HT_NOEXTERN, HT_NOFUNC, HT_NONE, HT_STRUCT };

// ── Section writing ─────────────────────────────────────────────────────

pub fn write_section(out: &mut Vec<u8>, id: u8, data: &[u8]) {
    out.push(id);
    write_leb128_u32(out, data.len() as u32);
    out.extend_from_slice(data);
}

pub fn write_name(out: &mut Vec<u8>, s: &str) {
    write_leb128_u32(out, s.len() as u32);
    out.extend_from_slice(s.as_bytes());
}

// ── LEB128 encoding ─────────────────────────────────────────────────────

pub fn write_leb128_u32(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

pub fn write_leb128_u64(out: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

pub fn write_leb128_i32(out: &mut Vec<u8>, mut value: i32) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 0x40 == 0) || (value == -1 && byte & 0x40 != 0);
        if !done {
            out.push(byte | 0x80);
        } else {
            out.push(byte);
            break;
        }
    }
}

pub fn write_leb128_i64(out: &mut Vec<u8>, mut value: i64) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 0x40 == 0) || (value == -1 && byte & 0x40 != 0);
        if !done {
            out.push(byte | 0x80);
        } else {
            out.push(byte);
            break;
        }
    }
}

pub fn encode_memarg_with_memidx(out: &mut Vec<u8>, align: u32, offset: u64, memidx: u32) {
    let encoded_align = if memidx == 0 { align } else { align | 0x40 };
    write_leb128_u32(out, encoded_align);
    write_leb128_u64(out, offset);
    if memidx != 0 {
        write_leb128_u32(out, memidx);
    }
}

// ── LEB128 decoding ─────────────────────────────────────────────────────

pub fn read_leb128_u32(data: &[u8]) -> (u32, usize) {
    let mut result = 0u32;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        result |= ((byte & 0x7f) as u32) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    (result, pos)
}

pub fn read_leb128_i32(data: &[u8]) -> (i32, usize) {
    let mut result = 0i64;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos] as i64;
        pos += 1;
        result |= (byte & 0x7f) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 64 && (byte & 0x40) != 0 {
                result |= !0i64 << shift;
            }
            break;
        }
    }
    (result as i32, pos)
}

pub fn read_leb128_i64(data: &[u8]) -> (i64, usize) {
    let mut result = 0i64;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos] as i64;
        pos += 1;
        result |= (byte & 0x7f) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 64 && (byte & 0x40) != 0 {
                result |= !0i64 << shift;
            }
            break;
        }
    }
    (result, pos)
}

pub fn skip_leb128(data: &[u8], pos: &mut usize) {
    while *pos < data.len() {
        let byte = data[*pos];
        *pos += 1;
        if byte & 0x80 == 0 {
            break;
        }
    }
}

// ── Bytecode operand helpers ────────────────────────────────────────────

pub fn read_u16(code: &[u8], ip: &mut usize) -> u16 {
    let hi = code[*ip] as u16;
    let lo = code[*ip + 1] as u16;
    *ip += 2;
    (hi << 8) | lo
}

pub fn read_i16(code: &[u8], ip: &mut usize) -> i16 {
    read_u16(code, ip) as i16
}

// ── Value serialization ─────────────────────────────────────────────────

pub fn encode_value(out: &mut Vec<u8>, val: &Value) {
    match val {
        Value::Null | Value::Undefined => out.push(0),
        Value::Bool(b) => {
            out.push(1);
            out.push(if *b { 1 } else { 0 });
        }
        Value::I32(n) => {
            out.push(2);
            out.extend_from_slice(&n.to_le_bytes());
        }
        Value::I64(n) => {
            out.push(3);
            out.extend_from_slice(&n.to_le_bytes());
        }
        Value::F64(n) => {
            out.push(4);
            out.extend_from_slice(&n.to_le_bytes());
        }
        Value::String(s) => {
            out.push(5);
            write_leb128_u32(out, s.len() as u32);
            out.extend_from_slice(s.as_bytes());
        }
        _ => out.push(0) }
}

pub fn decode_value(data: &[u8], pos: &mut usize) -> Value {
    if *pos >= data.len() {
        return Value::Null;
    }
    let tag = data[*pos];
    *pos += 1;
    match tag {
        0 => Value::Null,
        1 => {
            let b = data[*pos];
            *pos += 1;
            Value::Bool(b != 0)
        }
        2 => {
            let bytes: [u8; 4] = data[*pos..*pos + 4].try_into().unwrap_or([0; 4]);
            *pos += 4;
            Value::I32(i32::from_le_bytes(bytes))
        }
        3 => {
            let bytes: [u8; 8] = data[*pos..*pos + 8].try_into().unwrap_or([0; 8]);
            *pos += 8;
            Value::I64(i64::from_le_bytes(bytes))
        }
        4 => {
            let bytes: [u8; 8] = data[*pos..*pos + 8].try_into().unwrap_or([0; 8]);
            *pos += 8;
            Value::F64(f64::from_le_bytes(bytes))
        }
        5 => {
            let (len, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            let s = std::str::from_utf8(&data[*pos..*pos + len as usize]).unwrap_or("");
            *pos += len as usize;
            Value::String(Arc::from(s))
        }
        _ => Value::Null }
}
