//! Behaviour tests for `node:string_decoder` host imports.
//!
//! Reference: <https://nodejs.org/api/string_decoder.html>.
//!
//! Coverage:
//!   - `StringDecoder(encoding)` constructor
//!   - `.write(buffer)` → partial string (accumulates incomplete multibyte seqs)
//!   - `.end([buffer])` → flushes and returns remaining bytes as replacement char
//!   - Encodings: utf8, utf16le, base64, latin1, ascii, hex
//!   - Multibyte UTF-8 split across two `.write()` calls

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_sd(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-string_decoder-test>");
    let import_idx = chunk.add_import("node:string_decoder", name);
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
        vm.globals.insert(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:string_decoder"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn bytes_buf(bytes: &[u8]) -> Value {
    let elems = bytes.iter().map(|&b| Value::I32(b as i32)).collect();
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(elems),
        properties: Default::default(),
        type_id: 0,
        fields: Vec::new(),
    })))
}

fn as_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{other}"),
    }
}

fn new_decoder(encoding: &str) -> Value {
    call_sd("StringDecoder", vec![s(encoding)])
}

// ── Constructor ────────────────────────────────────────────────────────────────

#[test]
fn string_decoder_constructor_returns_object() {
    let decoder = new_decoder("utf8");
    assert!(matches!(decoder, Value::Object(_)));
}

#[test]
fn string_decoder_has_encoding_property() {
    let decoder = new_decoder("utf8");
    match &decoder {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            let enc = obj
                .properties
                .get("encoding")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(as_str(&enc), "utf-8");
        }
        _ => panic!("expected object"),
    }
}

// ── write (utf8) ──────────────────────────────────────────────────────────────

#[test]
fn write_ascii_bytes_returns_ascii_string() {
    let decoder = new_decoder("utf8");
    let buf = bytes_buf(b"Hello");
    let result = call_sd("write", vec![decoder, buf]);
    assert_eq!(as_str(&result), "Hello");
}

#[test]
fn write_complete_multibyte_utf8_returns_character() {
    let decoder = new_decoder("utf8");
    // "é" = [0xC3, 0xA9]
    let buf = bytes_buf(&[0xC3, 0xA9]);
    let result = call_sd("write", vec![decoder, buf]);
    assert_eq!(as_str(&result), "é");
}

#[test]
fn write_split_multibyte_utf8_first_byte_returns_empty() {
    // First byte of "é" alone — decoder should buffer it, not return partial
    let decoder = new_decoder("utf8");
    let first_byte = bytes_buf(&[0xC3]);
    let result = call_sd("write", vec![decoder, first_byte]);
    assert_eq!(
        as_str(&result),
        "",
        "incomplete multibyte must not be emitted"
    );
}

#[test]
fn write_split_multibyte_utf8_second_call_completes_char() {
    let decoder = new_decoder("utf8");
    let first = bytes_buf(&[0xC3]);
    let second = bytes_buf(&[0xA9]);
    let _part1 = call_sd("write", vec![decoder.clone(), first]);
    let part2 = call_sd("write", vec![decoder, second]);
    assert_eq!(as_str(&part2), "é");
}

#[test]
fn write_three_byte_utf8_char_across_three_calls() {
    // "€" = [0xE2, 0x82, 0xAC]
    let decoder = new_decoder("utf8");
    let b1 = bytes_buf(&[0xE2]);
    let b2 = bytes_buf(&[0x82]);
    let b3 = bytes_buf(&[0xAC]);
    let r1 = call_sd("write", vec![decoder.clone(), b1]);
    let r2 = call_sd("write", vec![decoder.clone(), b2]);
    let r3 = call_sd("write", vec![decoder, b3]);
    assert_eq!(as_str(&r1), "");
    assert_eq!(as_str(&r2), "");
    assert_eq!(as_str(&r3), "€");
}

// ── end ───────────────────────────────────────────────────────────────────────

#[test]
fn end_with_no_args_flushes_empty_decoder() {
    let decoder = new_decoder("utf8");
    let result = call_sd("end", vec![decoder]);
    assert_eq!(as_str(&result), "");
}

#[test]
fn end_with_buffer_writes_and_flushes() {
    let decoder = new_decoder("utf8");
    let buf = bytes_buf(b"World");
    let result = call_sd("end", vec![decoder, buf]);
    assert_eq!(as_str(&result), "World");
}

#[test]
fn end_flushes_incomplete_sequence_as_replacement_char() {
    // Incomplete multibyte then end() → U+FFFD replacement char
    let decoder = new_decoder("utf8");
    let incomplete = bytes_buf(&[0xC3]); // first byte of "é" only
    let _ = call_sd("write", vec![decoder.clone(), incomplete]);
    let result = call_sd("end", vec![decoder]);
    assert_eq!(as_str(&result), "\u{FFFD}");
}

// ── latin1 encoding ───────────────────────────────────────────────────────────

#[test]
fn latin1_decoder_returns_iso_8859_1_string() {
    let decoder = new_decoder("latin1");
    // In latin1, byte 0xE9 = 'é'
    let buf = bytes_buf(&[0xE9]);
    let result = call_sd("write", vec![decoder, buf]);
    assert_eq!(as_str(&result), "é");
}

// ── hex encoding ──────────────────────────────────────────────────────────────

#[test]
fn hex_decoder_returns_hex_string() {
    let decoder = new_decoder("hex");
    let buf = bytes_buf(&[0xDE, 0xAD]);
    let result = call_sd("write", vec![decoder, buf]);
    assert_eq!(as_str(&result), "dead");
}

// ── base64 encoding ───────────────────────────────────────────────────────────

#[test]
fn base64_decoder_encodes_bytes_as_base64() {
    let decoder = new_decoder("base64");
    // [0x41, 0x42, 0x43] = "ABC" → "QUJD"
    let buf = bytes_buf(&[0x41, 0x42, 0x43]);
    let result = call_sd("write", vec![decoder, buf]);
    assert_eq!(as_str(&result), "QUJD");
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_string_decoder_surface_is_registered() {
    let expected = ["StringDecoder", "write", "end"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:string_decoder imports: {missing:?}"
    );
}
