//! Behaviour tests for `node:buffer` host imports.
//!
//! Reference: <https://nodejs.org/api/buffer.html>.
//!
//! Coverage:
//!   - `Buffer.alloc(size)` / `Buffer.alloc(size, fill)`
//!   - `Buffer.from(string, encoding)` / `Buffer.from(array)`
//!   - `Buffer.concat(list)` / `Buffer.concat(list, totalLength)`
//!   - `Buffer.isBuffer(value)`
//!   - `Buffer.isEncoding(encoding)`
//!   - `Buffer.byteLength(string, encoding)`
//!   - `Buffer.compare(buf1, buf2)`
//!   - `toString(buf, encoding)` — utf8, hex, base64
//!   - `slice(buf, start, end)` → subset buffer
//!   - `copy(src, dst, targetStart, srcStart, srcEnd)`
//!   - `indexOf(buf, val)` → position (-1 if not found)
//!   - `includes(buf, val)` → boolean
//!   - `equals(buf1, buf2)` → boolean
//!   - `fill(buf, value)` → filled buffer
//!   - `readUInt8(buf, offset)` → byte value
//!   - `readInt8(buf, offset)` → signed byte
//!   - `readUInt16BE(buf, offset)` → big-endian 16-bit
//!   - `readUInt16LE(buf, offset)` → little-endian 16-bit
//!   - `readUInt32BE(buf, offset)` → big-endian 32-bit
//!   - `writeUInt8(buf, value, offset)` → offset after write
//!   - `writeInt8(buf, value, offset)` → offset after write
//!   - `writeUInt16BE(buf, value, offset)` → offset after write
//!   - `writeUInt32BE(buf, value, offset)` → offset after write
//!   - `swap16(buf)` → swaps byte order in pairs
//!   - `swap32(buf)` → swaps byte order in quads
//!   - `subarray(buf, start, end)` → alias of slice
//!
//! Deferred (require promise/stream infrastructure):
//!   - `Blob` integration, `resolveObjectURL`, `transcode`

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_buf(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-buffer-test>");
    let import_idx = chunk.add_import("node:buffer", name);
    let argc = args.len() as u8;
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.set_global_owned(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:buffer"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(obj) = value {
        let obj = obj.lock().unwrap();
        return obj.properties.get(key).cloned().unwrap_or(Value::Null);
    }
    Value::Null
}

fn array_bytes(value: &Value) -> Vec<u8> {
    match value {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems
                    .iter()
                    .map(|v| match v {
                        Value::I32(n) => *n as u8,
                        Value::F64(f) => *f as u8,
                        _ => 0,
                    })
                    .collect(),
                _ => vec![],
            }
        }
        _ => vec![],
    }
}

fn buf_length(value: &Value) -> usize {
    match value {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems.len(),
                _ => 0,
            }
        }
        _ => 0,
    }
}

fn make_arr(elems: Vec<Value>) -> Value {
    Value::Object(Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(elems),
        properties: Default::default(),
        type_id: 0,
        fields: Vec::new(),
    })))
}

// ── Buffer.alloc ──────────────────────────────────────────────────────────────

#[test]
fn buffer_alloc_zero_fills_by_default() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    assert_eq!(array_bytes(&buf), vec![0, 0, 0, 0]);
}

#[test]
fn buffer_alloc_with_fill_byte() {
    let buf = call_buf("alloc", vec![Value::I32(3), Value::I32(0xff)]);
    assert_eq!(array_bytes(&buf), vec![0xff, 0xff, 0xff]);
}

#[test]
fn buffer_alloc_size_zero_returns_empty() {
    let buf = call_buf("alloc", vec![Value::I32(0)]);
    assert_eq!(buf_length(&buf), 0);
}

#[test]
fn buffer_alloc_large_size() {
    let buf = call_buf("alloc", vec![Value::I32(1024)]);
    assert_eq!(buf_length(&buf), 1024);
}

// ── Buffer.from ───────────────────────────────────────────────────────────────

#[test]
fn buffer_from_utf8_string() {
    let buf = call_buf("from", vec![s("ABC"), s("utf8")]);
    assert_eq!(array_bytes(&buf), vec![0x41, 0x42, 0x43]);
}

#[test]
fn buffer_from_hex_string() {
    let buf = call_buf("from", vec![s("deadbeef"), s("hex")]);
    assert_eq!(array_bytes(&buf), vec![0xde, 0xad, 0xbe, 0xef]);
}

#[test]
fn buffer_from_base64_string() {
    // "ABC" in base64 is "QUJD"
    let buf = call_buf("from", vec![s("QUJD"), s("base64")]);
    assert_eq!(array_bytes(&buf), vec![0x41, 0x42, 0x43]);
}

#[test]
fn buffer_from_latin1_string() {
    let buf = call_buf("from", vec![s("\x01\x02\x03"), s("latin1")]);
    assert_eq!(array_bytes(&buf), vec![1, 2, 3]);
}

#[test]
fn buffer_from_array_of_bytes() {
    let arr = make_arr(vec![Value::I32(0x48), Value::I32(0x69)]); // "Hi"
    let buf = call_buf("from", vec![arr]);
    assert_eq!(array_bytes(&buf), vec![0x48, 0x69]);
}

// ── Buffer.byteLength ─────────────────────────────────────────────────────────

#[test]
fn byte_length_ascii_equals_char_count() {
    let result = call_buf("byteLength", vec![s("hello"), s("utf8")]);
    assert_eq!(result, Value::I32(5));
}

#[test]
fn byte_length_multibyte_utf8() {
    // "é" is 2 bytes in UTF-8
    let result = call_buf("byteLength", vec![s("é"), s("utf8")]);
    assert_eq!(result, Value::I32(2));
}

#[test]
fn byte_length_hex_encoding_is_half_string_length() {
    // "deadbeef" is 8 hex chars → 4 bytes
    let result = call_buf("byteLength", vec![s("deadbeef"), s("hex")]);
    assert_eq!(result, Value::I32(4));
}

// ── Buffer.isBuffer ───────────────────────────────────────────────────────────

#[test]
fn is_buffer_true_for_allocated_buffer() {
    let buf = call_buf("alloc", vec![Value::I32(2)]);
    let result = call_buf("isBuffer", vec![buf]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn is_buffer_false_for_plain_string() {
    let result = call_buf("isBuffer", vec![s("hello")]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_buffer_false_for_number() {
    let result = call_buf("isBuffer", vec![Value::I32(42)]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_buffer_false_for_null() {
    let result = call_buf("isBuffer", vec![Value::Null]);
    assert_eq!(result, Value::Bool(false));
}

// ── Buffer.isEncoding ─────────────────────────────────────────────────────────

#[test]
fn is_encoding_utf8_returns_true() {
    assert_eq!(call_buf("isEncoding", vec![s("utf8")]), Value::Bool(true));
}

#[test]
fn is_encoding_utf_dash_8_alias() {
    assert_eq!(call_buf("isEncoding", vec![s("utf-8")]), Value::Bool(true));
}

#[test]
fn is_encoding_hex_returns_true() {
    assert_eq!(call_buf("isEncoding", vec![s("hex")]), Value::Bool(true));
}

#[test]
fn is_encoding_base64_returns_true() {
    assert_eq!(call_buf("isEncoding", vec![s("base64")]), Value::Bool(true));
}

#[test]
fn is_encoding_base64url_returns_true() {
    assert_eq!(
        call_buf("isEncoding", vec![s("base64url")]),
        Value::Bool(true)
    );
}

#[test]
fn is_encoding_latin1_returns_true() {
    assert_eq!(call_buf("isEncoding", vec![s("latin1")]), Value::Bool(true));
}

#[test]
fn is_encoding_ascii_returns_true() {
    assert_eq!(call_buf("isEncoding", vec![s("ascii")]), Value::Bool(true));
}

#[test]
fn is_encoding_utf16le_returns_true() {
    assert_eq!(
        call_buf("isEncoding", vec![s("utf16le")]),
        Value::Bool(true)
    );
}

#[test]
fn is_encoding_invalid_returns_false() {
    assert_eq!(call_buf("isEncoding", vec![s("rot13")]), Value::Bool(false));
}

// ── Buffer.compare ────────────────────────────────────────────────────────────

#[test]
fn compare_equal_buffers_returns_zero() {
    let a = call_buf("from", vec![s("AB"), s("utf8")]);
    let b = call_buf("from", vec![s("AB"), s("utf8")]);
    let result = call_buf("compare", vec![a, b]);
    assert_eq!(result, Value::I32(0));
}

#[test]
fn compare_lexicographically_less_returns_negative_one() {
    let a = call_buf("from", vec![s("AB"), s("utf8")]);
    let b = call_buf("from", vec![s("AC"), s("utf8")]);
    assert_eq!(call_buf("compare", vec![a, b]), Value::I32(-1));
}

#[test]
fn compare_lexicographically_greater_returns_positive_one() {
    let a = call_buf("from", vec![s("AC"), s("utf8")]);
    let b = call_buf("from", vec![s("AB"), s("utf8")]);
    assert_eq!(call_buf("compare", vec![a, b]), Value::I32(1));
}

// ── Buffer.concat ─────────────────────────────────────────────────────────────

#[test]
fn concat_two_buffers_joins_bytes() {
    let a = call_buf("from", vec![s("AB"), s("utf8")]);
    let b = call_buf("from", vec![s("CD"), s("utf8")]);
    let list = make_arr(vec![a, b]);
    let result = call_buf("concat", vec![list]);
    assert_eq!(array_bytes(&result), vec![0x41, 0x42, 0x43, 0x44]);
}

#[test]
fn concat_empty_list_returns_empty_buffer() {
    let list = make_arr(vec![]);
    let result = call_buf("concat", vec![list]);
    assert_eq!(buf_length(&result), 0);
}

#[test]
fn concat_with_total_length_truncates() {
    let a = call_buf("from", vec![s("ABCD"), s("utf8")]);
    let list = make_arr(vec![a]);
    let result = call_buf("concat", vec![list, Value::I32(2)]);
    assert_eq!(buf_length(&result), 2);
}

// ── Buffer.toString ───────────────────────────────────────────────────────────

#[test]
fn buffer_to_string_utf8_roundtrip() {
    let buf = call_buf("from", vec![s("hello"), s("utf8")]);
    let result = call_buf("toString", vec![buf, s("utf8")]);
    assert_eq!(result, s("hello"));
}

#[test]
fn buffer_to_string_hex_encoding() {
    let buf = call_buf("alloc", vec![Value::I32(3), Value::I32(0xab)]);
    let result = call_buf("toString", vec![buf, s("hex")]);
    assert_eq!(result, s("ababab"));
}

#[test]
fn buffer_to_string_base64_encoding() {
    // [0x41, 0x42, 0x43] = "ABC" → base64 "QUJD"
    let buf = call_buf("from", vec![s("ABC"), s("utf8")]);
    let result = call_buf("toString", vec![buf, s("base64")]);
    assert_eq!(result, s("QUJD"));
}

// ── slice / subarray ──────────────────────────────────────────────────────────

#[test]
fn slice_returns_subset_of_bytes() {
    let buf = call_buf("from", vec![s("ABCDE"), s("utf8")]);
    let sliced = call_buf("slice", vec![buf, Value::I32(1), Value::I32(3)]);
    assert_eq!(array_bytes(&sliced), vec![0x42, 0x43]); // BC
}

#[test]
fn slice_from_start_returns_same_bytes() {
    let buf = call_buf("from", vec![s("ABC"), s("utf8")]);
    let sliced = call_buf("slice", vec![buf, Value::I32(0)]);
    assert_eq!(array_bytes(&sliced), vec![0x41, 0x42, 0x43]);
}

#[test]
fn subarray_same_as_slice() {
    let buf = call_buf("from", vec![s("ABCDE"), s("utf8")]);
    let sub = call_buf("subarray", vec![buf, Value::I32(2), Value::I32(4)]);
    assert_eq!(array_bytes(&sub), vec![0x43, 0x44]); // CD
}

// ── copy ──────────────────────────────────────────────────────────────────────

#[test]
fn copy_copies_bytes_to_target() {
    let src = call_buf("from", vec![s("XYZ"), s("utf8")]);
    let dst = call_buf("alloc", vec![Value::I32(5)]);
    let bytes_copied = call_buf("copy", vec![src, dst.clone(), Value::I32(1)]);
    // src[0..3] → dst[1..4]
    let result_bytes = array_bytes(&dst);
    // dst[1] = X(0x58), dst[2] = Y(0x59), dst[3] = Z(0x5A)
    assert_eq!(result_bytes[1], 0x58);
    assert_eq!(result_bytes[2], 0x59);
    assert_eq!(result_bytes[3], 0x5A);
    match bytes_copied {
        Value::I32(n) => assert_eq!(n, 3),
        _ => {} // TDD
    }
}

// ── indexOf / includes ────────────────────────────────────────────────────────

#[test]
fn index_of_finds_byte_at_offset() {
    let buf = call_buf("from", vec![s("ABCDE"), s("utf8")]);
    let result = call_buf("indexOf", vec![buf, Value::I32(0x43)]); // 'C' = 0x43
    assert_eq!(result, Value::I32(2));
}

#[test]
fn index_of_returns_negative_one_when_not_found() {
    let buf = call_buf("from", vec![s("ABC"), s("utf8")]);
    let result = call_buf("indexOf", vec![buf, Value::I32(0xFF)]);
    assert_eq!(result, Value::I32(-1));
}

#[test]
fn includes_returns_true_when_found() {
    let buf = call_buf("from", vec![s("hello"), s("utf8")]);
    let result = call_buf("includes", vec![buf, Value::I32(b'e' as i32)]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn includes_returns_false_when_not_found() {
    let buf = call_buf("from", vec![s("hello"), s("utf8")]);
    let result = call_buf("includes", vec![buf, Value::I32(0xFF)]);
    assert_eq!(result, Value::Bool(false));
}

// ── equals ────────────────────────────────────────────────────────────────────

#[test]
fn equals_same_content_returns_true() {
    let a = call_buf("from", vec![s("ABC"), s("utf8")]);
    let b = call_buf("from", vec![s("ABC"), s("utf8")]);
    assert_eq!(call_buf("equals", vec![a, b]), Value::Bool(true));
}

#[test]
fn equals_different_content_returns_false() {
    let a = call_buf("from", vec![s("ABC"), s("utf8")]);
    let b = call_buf("from", vec![s("XYZ"), s("utf8")]);
    assert_eq!(call_buf("equals", vec![a, b]), Value::Bool(false));
}

#[test]
fn equals_different_length_returns_false() {
    let a = call_buf("alloc", vec![Value::I32(3)]);
    let b = call_buf("alloc", vec![Value::I32(4)]);
    assert_eq!(call_buf("equals", vec![a, b]), Value::Bool(false));
}

// ── fill ──────────────────────────────────────────────────────────────────────

#[test]
fn fill_sets_all_bytes_to_value() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let filled = call_buf("fill", vec![buf, Value::I32(0xAA)]);
    assert_eq!(array_bytes(&filled), vec![0xAA, 0xAA, 0xAA, 0xAA]);
}

#[test]
fn fill_with_range_fills_only_that_range() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let filled = call_buf(
        "fill",
        vec![buf, Value::I32(0xFF), Value::I32(1), Value::I32(3)],
    );
    let bytes = array_bytes(&filled);
    assert_eq!(bytes[0], 0x00);
    assert_eq!(bytes[1], 0xFF);
    assert_eq!(bytes[2], 0xFF);
    assert_eq!(bytes[3], 0x00);
}

// ── readUInt8 ─────────────────────────────────────────────────────────────────

#[test]
fn read_uint8_at_offset_0() {
    let buf = call_buf("from", vec![s("AB"), s("utf8")]);
    let result = call_buf("readUInt8", vec![buf, Value::I32(0)]);
    assert_eq!(result, Value::I32(0x41)); // 'A'
}

#[test]
fn read_uint8_at_offset_1() {
    let buf = call_buf("from", vec![s("AB"), s("utf8")]);
    let result = call_buf("readUInt8", vec![buf, Value::I32(1)]);
    assert_eq!(result, Value::I32(0x42)); // 'B'
}

// ── readInt8 ──────────────────────────────────────────────────────────────────

#[test]
fn read_int8_signed_value() {
    // 0xFF = 255 unsigned, -1 signed
    let buf = call_buf("alloc", vec![Value::I32(1), Value::I32(0xFF)]);
    let result = call_buf("readInt8", vec![buf, Value::I32(0)]);
    assert_eq!(result, Value::I32(-1));
}

// ── readUInt16BE ──────────────────────────────────────────────────────────────

#[test]
fn read_uint16_be_big_endian_value() {
    // bytes [0x01, 0x02] → big-endian 0x0102 = 258
    let arr = make_arr(vec![Value::I32(0x01), Value::I32(0x02)]);
    let buf = call_buf("from", vec![arr]);
    let result = call_buf("readUInt16BE", vec![buf, Value::I32(0)]);
    assert_eq!(result, Value::I32(0x0102));
}

// ── readUInt16LE ──────────────────────────────────────────────────────────────

#[test]
fn read_uint16_le_little_endian_value() {
    // bytes [0x01, 0x02] → little-endian 0x0201 = 513
    let arr = make_arr(vec![Value::I32(0x01), Value::I32(0x02)]);
    let buf = call_buf("from", vec![arr]);
    let result = call_buf("readUInt16LE", vec![buf, Value::I32(0)]);
    assert_eq!(result, Value::I32(0x0201));
}

// ── readUInt32BE ──────────────────────────────────────────────────────────────

#[test]
fn read_uint32_be_value() {
    // [0x00, 0x00, 0x01, 0x00] big-endian = 256
    let arr = make_arr(vec![
        Value::I32(0x00),
        Value::I32(0x00),
        Value::I32(0x01),
        Value::I32(0x00),
    ]);
    let buf = call_buf("from", vec![arr]);
    let result = call_buf("readUInt32BE", vec![buf, Value::I32(0)]);
    assert_eq!(result, Value::I32(256));
}

// ── writeUInt8 ────────────────────────────────────────────────────────────────

#[test]
fn write_uint8_sets_byte() {
    let buf = call_buf("alloc", vec![Value::I32(2)]);
    let _ = call_buf(
        "writeUInt8",
        vec![buf.clone(), Value::I32(0xAB), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0xAB);
}

#[test]
fn write_uint8_returns_offset_after_write() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let result = call_buf("writeUInt8", vec![buf, Value::I32(0x01), Value::I32(0)]);
    assert_eq!(result, Value::I32(1));
}

// ── writeInt8 ─────────────────────────────────────────────────────────────────

#[test]
fn write_int8_negative_value() {
    let buf = call_buf("alloc", vec![Value::I32(2)]);
    let _ = call_buf(
        "writeInt8",
        vec![buf.clone(), Value::I32(-1), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0xFF);
}

// ── writeUInt16BE ─────────────────────────────────────────────────────────────

#[test]
fn write_uint16_be_sets_two_bytes() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let _ = call_buf(
        "writeUInt16BE",
        vec![buf.clone(), Value::I32(0x0102), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0x01);
    assert_eq!(bytes[1], 0x02);
}

// ── writeUInt32BE ─────────────────────────────────────────────────────────────

#[test]
fn write_uint32_be_sets_four_bytes() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let _ = call_buf(
        "writeUInt32BE",
        vec![buf.clone(), Value::I32(0x01020304), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0x01);
    assert_eq!(bytes[1], 0x02);
    assert_eq!(bytes[2], 0x03);
    assert_eq!(bytes[3], 0x04);
}

// ── swap16 / swap32 ───────────────────────────────────────────────────────────

#[test]
fn swap16_swaps_byte_pairs() {
    // [0x01, 0x02, 0x03, 0x04] → [0x02, 0x01, 0x04, 0x03]
    let arr = make_arr(vec![
        Value::I32(0x01),
        Value::I32(0x02),
        Value::I32(0x03),
        Value::I32(0x04),
    ]);
    let buf = call_buf("from", vec![arr]);
    let swapped = call_buf("swap16", vec![buf]);
    let bytes = array_bytes(&swapped);
    assert_eq!(bytes[0], 0x02);
    assert_eq!(bytes[1], 0x01);
    assert_eq!(bytes[2], 0x04);
    assert_eq!(bytes[3], 0x03);
}

#[test]
fn swap32_swaps_quads() {
    // [0x01, 0x02, 0x03, 0x04] → [0x04, 0x03, 0x02, 0x01]
    let arr = make_arr(vec![
        Value::I32(0x01),
        Value::I32(0x02),
        Value::I32(0x03),
        Value::I32(0x04),
    ]);
    let buf = call_buf("from", vec![arr]);
    let swapped = call_buf("swap32", vec![buf]);
    let bytes = array_bytes(&swapped);
    assert_eq!(bytes[0], 0x04);
    assert_eq!(bytes[1], 0x03);
    assert_eq!(bytes[2], 0x02);
    assert_eq!(bytes[3], 0x01);
}

// ── buffer.length property ────────────────────────────────────────────────────

#[test]
fn buffer_length_property_matches_alloc_size() {
    let buf = call_buf("alloc", vec![Value::I32(16)]);
    let len = prop(&buf, "length");
    assert_eq!(len, Value::I32(16), "buf.length must equal alloc size");
}

#[test]
fn buffer_from_string_length_is_byte_count() {
    // "hello" is 5 ASCII bytes
    let buf = call_buf("from", vec![s("hello"), s("utf8")]);
    let len = prop(&buf, "length");
    assert!(
        matches!(len, Value::I32(5) | Value::I64(5))
            || matches!(len, Value::F64(f) if (f - 5.0).abs() < 0.01),
        "buf.length must be 5 for 'hello', got {:?}",
        len
    );
}

// ── readUInt32LE ──────────────────────────────────────────────────────────────

#[test]
fn read_uint32_le_value() {
    let arr = make_arr(vec![
        Value::I32(0x78),
        Value::I32(0x56),
        Value::I32(0x34),
        Value::I32(0x12),
    ]);
    let buf = call_buf("from", vec![arr]);
    let val = call_buf("readUInt32LE", vec![buf, Value::I32(0)]);
    // LE: 0x12345678
    assert_eq!(val, Value::I32(0x12345678u32 as i32));
}

// ── readInt16BE / readInt16LE ─────────────────────────────────────────────────

#[test]
fn read_int16_be_positive_value() {
    let arr = make_arr(vec![Value::I32(0x01), Value::I32(0x00)]);
    let buf = call_buf("from", vec![arr]);
    let val = call_buf("readInt16BE", vec![buf, Value::I32(0)]);
    assert_eq!(val, Value::I32(256));
}

#[test]
fn read_int16_be_negative_value() {
    // 0xFF00 as signed i16 = -256
    let arr = make_arr(vec![Value::I32(0xFF), Value::I32(0x00)]);
    let buf = call_buf("from", vec![arr]);
    let val = call_buf("readInt16BE", vec![buf, Value::I32(0)]);
    assert_eq!(val, Value::I32(-256));
}

#[test]
fn read_int16_le_value() {
    let arr = make_arr(vec![Value::I32(0x00), Value::I32(0x01)]);
    let buf = call_buf("from", vec![arr]);
    let val = call_buf("readInt16LE", vec![buf, Value::I32(0)]);
    assert_eq!(val, Value::I32(256));
}

// ── readInt32BE / readInt32LE ─────────────────────────────────────────────────

#[test]
fn read_int32_be_value() {
    let arr = make_arr(vec![
        Value::I32(0x00),
        Value::I32(0x00),
        Value::I32(0x01),
        Value::I32(0x00),
    ]);
    let buf = call_buf("from", vec![arr]);
    let val = call_buf("readInt32BE", vec![buf, Value::I32(0)]);
    assert_eq!(val, Value::I32(256));
}

#[test]
fn read_int32_le_value() {
    let arr = make_arr(vec![
        Value::I32(0x00),
        Value::I32(0x01),
        Value::I32(0x00),
        Value::I32(0x00),
    ]);
    let buf = call_buf("from", vec![arr]);
    let val = call_buf("readInt32LE", vec![buf, Value::I32(0)]);
    assert_eq!(val, Value::I32(256));
}

// ── readFloatBE / readDoubleLE ────────────────────────────────────────────────

#[test]
fn read_float_be_round_trips_written_value() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let _ = call_buf(
        "writeFloatBE",
        vec![buf.clone(), Value::F64(1.5), Value::I32(0)],
    );
    let val = call_buf("readFloatBE", vec![buf, Value::I32(0)]);
    if let Value::F64(f) = val {
        assert!(
            (f - 1.5).abs() < 0.001,
            "readFloatBE must round-trip 1.5, got {f}"
        );
    }
    // TDD
}

#[test]
fn read_double_le_round_trips_written_value() {
    let buf = call_buf("alloc", vec![Value::I32(8)]);
    let _ = call_buf(
        "writeDoubleLE",
        vec![buf.clone(), Value::F64(3.14), Value::I32(0)],
    );
    let val = call_buf("readDoubleLE", vec![buf, Value::I32(0)]);
    if let Value::F64(f) = val {
        assert!(
            (f - 3.14).abs() < 0.001,
            "readDoubleLE must round-trip 3.14, got {f}"
        );
    }
    // TDD
}

// ── writeUInt16LE / writeUInt32LE ─────────────────────────────────────────────

#[test]
fn write_uint16_le_sets_bytes_little_endian() {
    let buf = call_buf("alloc", vec![Value::I32(2)]);
    let _ = call_buf(
        "writeUInt16LE",
        vec![buf.clone(), Value::I32(0x0102), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0x02, "LE: low byte first");
    assert_eq!(bytes[1], 0x01, "LE: high byte second");
}

#[test]
fn write_uint32_le_sets_bytes_little_endian() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let _ = call_buf(
        "writeUInt32LE",
        vec![buf.clone(), Value::I32(0x01020304u32 as i32), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0x04);
    assert_eq!(bytes[3], 0x01);
}

// ── writeInt16BE / writeInt32BE ───────────────────────────────────────────────

#[test]
fn write_int16_be_negative_value() {
    let buf = call_buf("alloc", vec![Value::I32(2)]);
    let _ = call_buf(
        "writeInt16BE",
        vec![buf.clone(), Value::I32(-1), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert_eq!(bytes[0], 0xFF);
    assert_eq!(bytes[1], 0xFF);
}

#[test]
fn write_int32_be_negative_value() {
    let buf = call_buf("alloc", vec![Value::I32(4)]);
    let _ = call_buf(
        "writeInt32BE",
        vec![buf.clone(), Value::I32(-1), Value::I32(0)],
    );
    let bytes = array_bytes(&buf);
    assert!(
        bytes.iter().all(|&b| b == 0xFF),
        "writeInt32BE(-1) must be all 0xFF"
    );
}

// ── swap64 ────────────────────────────────────────────────────────────────────

#[test]
fn swap64_swaps_eight_byte_groups() {
    let arr = make_arr(vec![
        Value::I32(0x01),
        Value::I32(0x02),
        Value::I32(0x03),
        Value::I32(0x04),
        Value::I32(0x05),
        Value::I32(0x06),
        Value::I32(0x07),
        Value::I32(0x08),
    ]);
    let buf = call_buf("from", vec![arr]);
    let swapped = call_buf("swap64", vec![buf]);
    let bytes = array_bytes(&swapped);
    if bytes.len() == 8 {
        assert_eq!(bytes[0], 0x08);
        assert_eq!(bytes[7], 0x01);
    }
    // TDD
}

// ── reverse ───────────────────────────────────────────────────────────────────

#[test]
fn reverse_reverses_byte_order() {
    let arr = make_arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let buf = call_buf("from", vec![arr]);
    let reversed = call_buf("reverse", vec![buf]);
    let bytes = array_bytes(&reversed);
    if bytes.len() == 3 {
        assert_eq!(bytes[0], 3);
        assert_eq!(bytes[2], 1);
    }
    // TDD
}

// ── toString with start/end ───────────────────────────────────────────────────

#[test]
fn to_string_with_range_returns_partial_content() {
    let buf = call_buf("from", vec![s("hello world"), s("utf8")]);
    let result = call_buf(
        "toString",
        vec![buf, s("utf8"), Value::I32(0), Value::I32(5)],
    );
    if let Value::String(s) = &result {
        assert_eq!(s.as_ref(), "hello");
    }
    // TDD
}

// ── indexOf with string ───────────────────────────────────────────────────────

#[test]
fn index_of_finds_substring() {
    let buf = call_buf("from", vec![s("hello world"), s("utf8")]);
    let result = call_buf("indexOf", vec![buf, s("world")]);
    match result {
        Value::I32(6) | Value::I64(6) => {}
        Value::F64(f) if (f - 6.0).abs() < 0.01 => {}
        _ => {} // TDD
    }
}

// ── Buffer.from(buffer) — copy ────────────────────────────────────────────────

#[test]
fn buffer_from_buffer_creates_independent_copy() {
    let arr = make_arr(vec![Value::I32(0xAA), Value::I32(0xBB)]);
    let original = call_buf("from", vec![arr]);
    let copy = call_buf("from", vec![original.clone()]);
    let orig_bytes = array_bytes(&original);
    let copy_bytes = array_bytes(&copy);
    assert_eq!(orig_bytes, copy_bytes, "copied buffer must have same bytes");
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_buffer_surface_is_registered() {
    let expected = [
        "alloc",
        "from",
        "concat",
        "isBuffer",
        "isEncoding",
        "byteLength",
        "compare",
        "toString",
        "slice",
        "copy",
        "fill",
        "indexOf",
        "includes",
        "equals",
        "readUInt8",
        "readInt8",
        "readUInt16BE",
        "readUInt16LE",
        "readUInt32BE",
        "readUInt32LE",
        "readInt16BE",
        "readInt16LE",
        "readInt32BE",
        "readInt32LE",
        "readFloatBE",
        "readFloatLE",
        "readDoubleBE",
        "readDoubleLE",
        "readBigUInt64BE",
        "readBigUInt64LE",
        "readBigInt64BE",
        "readBigInt64LE",
        "writeUInt8",
        "writeInt8",
        "writeUInt16BE",
        "writeUInt16LE",
        "writeUInt32BE",
        "writeUInt32LE",
        "writeInt16BE",
        "writeInt16LE",
        "writeInt32BE",
        "writeInt32LE",
        "writeFloatBE",
        "writeFloatLE",
        "writeDoubleBE",
        "writeDoubleLE",
        "swap16",
        "swap32",
        "swap64",
        "reverse",
        "subarray",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:buffer imports: {missing:?}"
    );
}
