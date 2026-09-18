//! WASM binary encoding/decoding primitives.
//! Constants, LEB128, section writing, value serialization.

use std::collections::{HashMap, HashSet};
use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind, Value};

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
pub const SECTION_DATA: u8 = 11;
pub const SECTION_DATA_COUNT: u8 = 12;
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
    HT_ANY, HT_ARRAY, HT_EQ, HT_EXTERN, HT_FUNC, HT_I31, HT_NOEXTERN, HT_NOFUNC, HT_NONE,
    HT_STRUCT, HeapType,
};

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
    let mut shift = 0u32;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        // ⛔ A u32 LEB IS AT MOST FIVE BYTES. Past that the shift overflows and
        // this PANICKED — `binary-leb128.wast` feeds overlong encodings on
        // purpose, so the decoder has to reject them, not abort the process.
        // `0` bytes read is the failure signal every caller already tests.
        if shift >= 32 {
            return (0, 0);
        }
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

// ── Heaptype encoding ──────────────────────────────────────────────────

/// Write a heaptype immediate.
///
/// Per the GC proposal this is an `s33`: abstract heap types are negative
/// values and concrete heap types are non-negative type indices. This is not
/// the same encoding as a plain `typeidx`, which is an unsigned LEB.
pub fn write_heaptype(out: &mut Vec<u8>, heaptype: HeapType) {
    write_leb128_i32(out, heaptype.to_sleb());
}

pub fn write_concrete_heaptype(out: &mut Vec<u8>, type_idx: u32) {
    write_heaptype(out, HeapType::Concrete(type_idx));
}

pub fn encode_concrete_heaptype(type_idx: u32) -> Vec<u8> {
    let mut out = Vec::new();
    write_concrete_heaptype(&mut out, type_idx);
    out
}

/// Write Custom Descriptors' exact heaptype form: `0x62 x:u32`.
pub fn write_exact_heaptype(out: &mut Vec<u8>, type_idx: u32) {
    out.push(HEAPTYPE_EXACT);
    write_leb128_u32(out, type_idx);
}

/// Read a heaptype immediate, including Custom Descriptors' exact form.
///
/// Exactness narrows the reference type but still names the same concrete
/// type index, so the returned `HeapType` is concrete.
pub fn read_heaptype_immediate(data: &[u8], pos: &mut usize) -> HeapType {
    if data.get(*pos) == Some(&HEAPTYPE_EXACT) {
        *pos += 1;
        let (index, read) = read_leb128_u32(&data[*pos..]);
        *pos += read;
        return HeapType::Concrete(index);
    }
    let (value, read) = read_leb128_i32(&data[*pos..]);
    *pos += read;
    HeapType::from_sleb(value)
}

pub fn skip_heaptype_immediate(data: &[u8], pos: &mut usize) {
    if data.get(*pos) == Some(&HEAPTYPE_EXACT) {
        *pos += 1;
        skip_leb128(data, pos);
        return;
    }
    skip_leb128(data, pos);
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

/// A fixed-width big-endian `u32` bytecode operand (`OperandFormat::U32`).
pub fn read_u32_be(code: &[u8], ip: &mut usize) -> u32 {
    let mut v = 0u32;
    for _ in 0..4 {
        v = (v << 8) | code[*ip] as u32;
        *ip += 1;
    }
    v
}

// ── Value serialization ─────────────────────────────────────────────────

const MAX_SERIALIZED_VALUE_GRAPH_DEPTH: usize = 128;
const MAX_SERIALIZED_VALUE_GRAPH_VALUES: usize = 16_384;

pub fn encode_value(out: &mut Vec<u8>, val: &Value) {
    let mut ctx = ValueEncodeContext::default();
    encode_value_with_context(out, val, &mut ctx);
}

pub struct ValueEncodeContext {
    object_ids: HashMap<usize, u32>,
}

impl Default for ValueEncodeContext {
    fn default() -> Self {
        Self {
            object_ids: HashMap::new(),
        }
    }
}

pub fn encode_value_with_context(out: &mut Vec<u8>, val: &Value, ctx: &mut ValueEncodeContext) {
    let mut seen = HashSet::new();
    encode_value_inner(out, val, &mut seen, ctx);
}

fn encode_value_inner(
    out: &mut Vec<u8>,
    val: &Value,
    seen: &mut HashSet<usize>,
    ctx: &mut ValueEncodeContext,
) {
    match val {
        Value::Null => out.push(0),
        Value::Undefined => out.push(12),
        Value::TypedNull(type_id) => {
            out.push(13);
            write_leb128_u32(out, *type_id as u32);
        }
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
        Value::F32(n) => {
            out.push(14);
            out.extend_from_slice(&n.to_le_bytes());
        }
        Value::String(s) => {
            out.push(5);
            write_leb128_u32(out, s.len() as u32);
            out.extend_from_slice(s.as_bytes());
        }
        Value::Symbol(s) => {
            out.push(15);
            write_leb128_u32(out, s.len() as u32);
            out.extend_from_slice(s.as_bytes());
        }
        Value::V128(bytes) => {
            out.push(16);
            out.extend_from_slice(bytes);
        }
        Value::Object(obj) => encode_object(out, obj, seen, ctx),
        _ => out.push(0),
    }
}

fn encode_object(
    out: &mut Vec<u8>,
    obj: &Arc<Mutex<Object>>,
    seen: &mut HashSet<usize>,
    ctx: &mut ValueEncodeContext,
) {
    let ptr = Arc::as_ptr(obj) as usize;
    if let Some(id) = ctx.object_ids.get(&ptr) {
        out.push(17);
        write_leb128_u32(out, *id);
        return;
    }
    let mut check_seen = HashSet::new();
    let mut budget = MAX_SERIALIZED_VALUE_GRAPH_VALUES;
    if !is_serializable_data_object(obj, &mut check_seen, 0, &mut budget) {
        out.push(0);
        return;
    }

    if !seen.insert(ptr) {
        out.push(0);
        return;
    }

    let id = ctx.object_ids.len() as u32;
    ctx.object_ids.insert(ptr, id);

    let Ok(guard) = obj.lock() else {
        seen.remove(&ptr);
        out.push(0);
        return;
    };

    match &guard.kind {
        ObjectKind::Ordinary => {
            out.push(6);
            write_leb128_u32(out, guard.type_id as u32);
            encode_properties(out, &guard.properties, seen, ctx);
            encode_value_vec(out, &guard.fields, seen, ctx);
        }
        ObjectKind::Array(elements) => {
            out.push(7);
            write_leb128_u32(out, guard.type_id as u32);
            encode_value_vec(out, elements, seen, ctx);
            encode_properties(out, &guard.properties, seen, ctx);
            encode_value_vec(out, &guard.fields, seen, ctx);
        }
        ObjectKind::Map(entries) => {
            out.push(8);
            write_leb128_u32(out, entries.len() as u32);
            for (key, value) in entries {
                encode_value_inner(out, key, seen, ctx);
                encode_value_inner(out, value, seen, ctx);
            }
        }
        ObjectKind::Set(entries) => {
            out.push(9);
            write_leb128_u32(out, entries.len() as u32);
            for value in entries {
                encode_value_inner(out, value, seen, ctx);
            }
        }
        _ => out.push(0),
    }

    seen.remove(&ptr);
}

fn is_serializable_data_value(
    val: &Value,
    seen: &mut HashSet<usize>,
    depth: usize,
    budget: &mut usize,
) -> bool {
    if depth > MAX_SERIALIZED_VALUE_GRAPH_DEPTH || *budget == 0 {
        return false;
    }
    *budget -= 1;
    match val {
        Value::Null
        | Value::Undefined
        | Value::TypedNull(_)
        | Value::Bool(_)
        | Value::I32(_)
        | Value::I64(_)
        | Value::F32(_)
        | Value::F64(_)
        | Value::String(_)
        | Value::Symbol(_)
        | Value::V128(_) => true,
        Value::Object(obj) => is_serializable_data_object(obj, seen, depth + 1, budget),
        _ => false,
    }
}

fn is_serializable_data_object(
    obj: &Arc<Mutex<Object>>,
    seen: &mut HashSet<usize>,
    depth: usize,
    budget: &mut usize,
) -> bool {
    if depth > MAX_SERIALIZED_VALUE_GRAPH_DEPTH || *budget == 0 {
        return false;
    }
    *budget -= 1;

    let ptr = Arc::as_ptr(obj) as usize;
    if !seen.insert(ptr) {
        return false;
    }

    let result = {
        let Ok(guard) = obj.lock() else {
            seen.remove(&ptr);
            return false;
        };

        let properties_ok = guard.properties.is_empty();
        let fields_ok = guard
            .fields
            .iter()
            .all(|value| is_serializable_data_value(value, seen, depth + 1, budget));

        properties_ok
            && fields_ok
            && match &guard.kind {
                ObjectKind::Ordinary => true,
                ObjectKind::Array(elements) => elements
                    .iter()
                    .all(|value| is_serializable_data_value(value, seen, depth + 1, budget)),
                ObjectKind::Map(entries) => entries.iter().all(|(key, value)| {
                    is_serializable_data_value(key, seen, depth + 1, budget)
                        && is_serializable_data_value(value, seen, depth + 1, budget)
                }),
                ObjectKind::Set(entries) => entries
                    .iter()
                    .all(|value| is_serializable_data_value(value, seen, depth + 1, budget)),
                ObjectKind::ArrayBuffer(_)
                | ObjectKind::TypedArray(_)
                | ObjectKind::Function(_)
                | ObjectKind::HostFunction(_)
                | ObjectKind::ModuleNamespace
                | ObjectKind::Continuation(_)
                | ObjectKind::Future { .. }
                | ObjectKind::Stream { .. } => false,
            }
    };

    seen.remove(&ptr);
    result
}

fn encode_properties(
    out: &mut Vec<u8>,
    properties: &indexmap::IndexMap<String, Value>,
    seen: &mut HashSet<usize>,
    ctx: &mut ValueEncodeContext,
) {
    write_leb128_u32(out, properties.len() as u32);
    for (key, value) in properties {
        write_name(out, key);
        encode_value_inner(out, value, seen, ctx);
    }
}

fn encode_value_vec(
    out: &mut Vec<u8>,
    values: &[Value],
    seen: &mut HashSet<usize>,
    ctx: &mut ValueEncodeContext,
) {
    write_leb128_u32(out, values.len() as u32);
    for value in values {
        encode_value_inner(out, value, seen, ctx);
    }
}

pub fn decode_value(data: &[u8], pos: &mut usize) -> Value {
    let mut ctx = ValueDecodeContext::default();
    decode_value_with_context(data, pos, &mut ctx)
}

#[derive(Default)]
pub struct ValueDecodeContext {
    objects: Vec<Value>,
}

pub fn decode_value_with_context(
    data: &[u8],
    pos: &mut usize,
    ctx: &mut ValueDecodeContext,
) -> Value {
    if *pos >= data.len() {
        return Value::Null;
    }
    let tag = data[*pos];
    *pos += 1;
    match tag {
        0 => Value::Null,
        12 => Value::Undefined,
        13 => {
            let (type_id, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            Value::TypedNull(type_id as usize)
        }
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
        6 => {
            let (type_id, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            let id = ctx.objects.len();
            ctx.objects.push(Value::Null);
            let properties = decode_properties(data, pos, ctx);
            let fields = decode_value_vec(data, pos, ctx);
            let value = Value::Object(Arc::new(Mutex::new(Object {
                properties,
                kind: ObjectKind::Ordinary,
                type_id: type_id as usize,
                fields,
            })));
            ctx.objects[id] = value.clone();
            value
        }
        7 => {
            let (type_id, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            let id = ctx.objects.len();
            ctx.objects.push(Value::Null);
            let elements = decode_value_vec(data, pos, ctx);
            let properties = decode_properties(data, pos, ctx);
            let fields = decode_value_vec(data, pos, ctx);
            let value = Value::Object(Arc::new(Mutex::new(Object {
                properties,
                kind: ObjectKind::Array(elements),
                type_id: type_id as usize,
                fields,
            })));
            ctx.objects[id] = value.clone();
            value
        }
        8 => {
            let (len, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            let id = ctx.objects.len();
            ctx.objects.push(Value::Null);
            let mut entries = indexmap::IndexMap::new();
            for _ in 0..len {
                let key = decode_value_with_context(data, pos, ctx);
                let value = decode_value_with_context(data, pos, ctx);
                entries.insert(key, value);
            }
            let value = Value::Object(Arc::new(Mutex::new(Object {
                properties: indexmap::IndexMap::new(),
                kind: ObjectKind::Map(entries),
                type_id: 0,
                fields: Vec::new(),
            })));
            ctx.objects[id] = value.clone();
            value
        }
        9 => {
            let (len, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            let id = ctx.objects.len();
            ctx.objects.push(Value::Null);
            let mut entries = indexmap::IndexSet::new();
            for _ in 0..len {
                entries.insert(decode_value_with_context(data, pos, ctx));
            }
            let value = Value::Object(Arc::new(Mutex::new(Object {
                properties: indexmap::IndexMap::new(),
                kind: ObjectKind::Set(entries),
                type_id: 0,
                fields: Vec::new(),
            })));
            ctx.objects[id] = value.clone();
            value
        }
        14 => {
            let bytes: [u8; 4] = data[*pos..*pos + 4].try_into().unwrap_or([0; 4]);
            *pos += 4;
            Value::F32(f32::from_le_bytes(bytes))
        }
        15 => {
            let (len, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            let s = std::str::from_utf8(&data[*pos..*pos + len as usize]).unwrap_or("");
            *pos += len as usize;
            Value::Symbol(Arc::from(s))
        }
        16 => {
            let mut bytes = [0u8; 16];
            if *pos + 16 <= data.len() {
                bytes.copy_from_slice(&data[*pos..*pos + 16]);
                *pos += 16;
            }
            Value::V128(bytes)
        }
        17 => {
            let (id, n) = read_leb128_u32(&data[*pos..]);
            *pos += n;
            ctx.objects
                .get(id as usize)
                .cloned()
                .unwrap_or(Value::Null)
        }
        _ => Value::Null,
    }
}

fn decode_properties(
    data: &[u8],
    pos: &mut usize,
    ctx: &mut ValueDecodeContext,
) -> indexmap::IndexMap<String, Value> {
    let (len, n) = read_leb128_u32(&data[*pos..]);
    *pos += n;
    let mut properties = indexmap::IndexMap::new();
    for _ in 0..len {
        let (key_len, read) = read_leb128_u32(&data[*pos..]);
        *pos += read;
        let key = std::str::from_utf8(&data[*pos..*pos + key_len as usize])
            .unwrap_or("")
            .to_string();
        *pos += key_len as usize;
        let value = decode_value_with_context(data, pos, ctx);
        properties.insert(key, value);
    }
    properties
}

fn decode_value_vec(data: &[u8], pos: &mut usize, ctx: &mut ValueDecodeContext) -> Vec<Value> {
    let (len, n) = read_leb128_u32(&data[*pos..]);
    *pos += n;
    let mut values = Vec::with_capacity(len as usize);
    for _ in 0..len {
        values.push(decode_value_with_context(data, pos, ctx));
    }
    values
}
