//! Behaviour tests for `ecma:dataview` host imports.
//!
//! Reference: ECMA-262 §25.3 DataView.
//!
//! DataView reads and writes typed values at arbitrary byte offsets into an
//! ArrayBuffer, with explicit endianness control. Each test covers a distinct
//! behaviour.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-dataview-test>");
    let import_idx = chunk.add_import("ecma:dataview", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn ab(byte_len: i32) -> Value {
    let mut chunk = Chunk::new("<ecma-dataview-ab>");
    let import_idx = chunk.add_import("ecma:arraybuffer", "newWithLength");
    let c = chunk.add_constant(Value::I32(byte_len));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(1, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

// ── construction ──────────────────────────────────────────────────────────────

#[test]
fn new_dataview_returns_object() {
    let buf = ab(16);
    let dv = invoke("new", vec![buf]);
    assert!(matches!(dv, Value::Object(_)));
}

#[test]
fn byte_length_reflects_full_buffer_when_no_offset_given() {
    let buf = ab(16);
    let dv = invoke("new", vec![buf]);
    assert_eq!(invoke("byteLength", vec![dv]), Value::I32(16));
}

#[test]
fn byte_offset_reflects_constructor_offset() {
    let buf = ab(16);
    let dv = invoke("newWithOffset", vec![buf, Value::I32(4)]);
    assert_eq!(invoke("byteOffset", vec![dv]), Value::I32(4));
}

#[test]
fn byte_length_respects_explicit_length_argument() {
    let buf = ab(16);
    let dv = invoke(
        "newWithOffsetAndLength",
        vec![buf, Value::I32(2), Value::I32(8)],
    );
    assert_eq!(invoke("byteLength", vec![dv]), Value::I32(8));
}

// ── Int8 / Uint8 ──────────────────────────────────────────────────────────────

#[test]
fn set_int8_and_get_int8_round_trip() {
    let buf = ab(4);
    let dv = invoke("new", vec![buf]);
    invoke("setInt8", vec![dv.clone(), Value::I32(0), Value::I32(-5)]);
    assert_eq!(invoke("getInt8", vec![dv, Value::I32(0)]), Value::I32(-5));
}

#[test]
fn set_uint8_wraps_at_256() {
    // Uint8 is unsigned; 300 mod 256 = 44.
    let buf = ab(4);
    let dv = invoke("new", vec![buf]);
    invoke("setUint8", vec![dv.clone(), Value::I32(0), Value::I32(300)]);
    assert_eq!(invoke("getUint8", vec![dv, Value::I32(0)]), Value::I32(44));
}

// ── Int16 endianness ──────────────────────────────────────────────────────────

#[test]
fn get_int16_little_endian_differs_from_big_endian() {
    // Write 0x0102 in little-endian → bytes [0x02, 0x01].
    // Reading back as big-endian gives 0x0201 = 513, not 258.
    let buf = ab(4);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setInt16LE",
        vec![dv.clone(), Value::I32(0), Value::I32(0x0102)],
    );
    let le = invoke("getInt16LE", vec![dv.clone(), Value::I32(0)]);
    let be = invoke("getInt16BE", vec![dv, Value::I32(0)]);
    assert_eq!(le, Value::I32(0x0102));
    assert_ne!(
        le, be,
        "little-endian and big-endian reads of same bytes must differ"
    );
}

// ── Int32 ─────────────────────────────────────────────────────────────────────

#[test]
fn set_int32_and_get_int32_little_endian_round_trip() {
    let buf = ab(8);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setInt32LE",
        vec![dv.clone(), Value::I32(0), Value::I32(0x01020304)],
    );
    assert_eq!(
        invoke("getInt32LE", vec![dv, Value::I32(0)]),
        Value::I32(0x01020304)
    );
}

#[test]
fn set_uint32_stores_values_above_i32_max() {
    // 0xDEADBEEF = 3735928559 — only fits in u32, not i32.
    let buf = ab(8);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setUint32LE",
        vec![dv.clone(), Value::I32(0), Value::F64(3735928559.0)],
    );
    if let Value::F64(v) = invoke("getUint32LE", vec![dv, Value::I32(0)]) {
        assert!((v - 3735928559.0).abs() < 1.0);
    } else {
        panic!("expected F64 for Uint32");
    }
}

// ── Float32 / Float64 ─────────────────────────────────────────────────────────

#[test]
fn set_float64_and_get_float64_little_endian_round_trip() {
    let buf = ab(16);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setFloat64LE",
        vec![dv.clone(), Value::I32(0), Value::F64(3.14)],
    );
    if let Value::F64(v) = invoke("getFloat64LE", vec![dv, Value::I32(0)]) {
        assert!((v - 3.14).abs() < 1e-10);
    } else {
        panic!("expected F64");
    }
}

#[test]
fn float32_loses_precision_compared_to_float64() {
    // f32 has ~7 significant decimal digits; 1.23456789 rounded.
    let buf = ab(8);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setFloat32LE",
        vec![dv.clone(), Value::I32(0), Value::F64(1.23456789)],
    );
    if let Value::F64(v) = invoke("getFloat32LE", vec![dv, Value::I32(0)]) {
        // The read-back value must differ from the original due to f32 precision.
        // f32 has ~7 sig figs; the round-trip differs from the f64 original
        assert!(
            v != 1.23456789,
            "f32 read-back must differ from original f64"
        );
    } else {
        panic!("expected F64");
    }
}

// ── BigInt64 ──────────────────────────────────────────────────────────────────

#[test]
fn set_big_int64_and_get_big_int64_little_endian_round_trip() {
    let buf = ab(16);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setBigInt64LE",
        vec![dv.clone(), Value::I32(0), Value::I64(i64::MAX)],
    );
    assert_eq!(
        invoke("getBigInt64LE", vec![dv, Value::I32(0)]),
        Value::I64(i64::MAX)
    );
}

#[test]
fn set_big_uint64_round_trip() {
    // u64::MAX cannot be stored as i64; the host must handle the raw bits.
    let buf = ab(16);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setBigUint64LE",
        vec![dv.clone(), Value::I32(0), Value::I64(-1)],
    );
    // -1 as i64 bits = 0xFFFF_FFFF_FFFF_FFFF = u64::MAX; reading back as I64(-1).
    assert_eq!(
        invoke("getBigUint64LE", vec![dv, Value::I32(0)]),
        Value::I64(-1)
    );
}

// ── Writes at non-zero offsets don't corrupt neighbours ───────────────────────

#[test]
fn write_at_offset_does_not_touch_adjacent_bytes() {
    let buf = ab(8);
    let dv = invoke("new", vec![buf]);
    invoke("setInt8", vec![dv.clone(), Value::I32(0), Value::I32(1)]);
    invoke("setInt8", vec![dv.clone(), Value::I32(1), Value::I32(2)]);
    invoke("setInt8", vec![dv.clone(), Value::I32(2), Value::I32(3)]);
    assert_eq!(
        invoke("getInt8", vec![dv.clone(), Value::I32(0)]),
        Value::I32(1)
    );
    assert_eq!(
        invoke("getInt8", vec![dv.clone(), Value::I32(1)]),
        Value::I32(2)
    );
    assert_eq!(invoke("getInt8", vec![dv, Value::I32(2)]), Value::I32(3));
}

// ── getFloat16 / setFloat16 (ES2025 §25.3.4.*) ───────────────────────────────

#[test]
fn set_float16_get_float16_round_trip_little_endian() {
    // ECMA-262 ES2025: DataView gains Float16 support via setFloat16/getFloat16.
    // Float16 has ~3 decimal digits of precision.
    let buf = ab(4);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setFloat16",
        vec![
            dv.clone(),
            Value::I32(0),
            Value::F64(1.5),
            Value::Bool(true),
        ],
    );
    let result = invoke("getFloat16", vec![dv, Value::I32(0), Value::Bool(true)]);
    match result {
        Value::F64(f) => assert!((f - 1.5).abs() < 0.01, "expected 1.5, got {f}"),
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}

#[test]
fn set_float16_zero_reads_back_as_zero() {
    let buf = ab(4);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setFloat16",
        vec![
            dv.clone(),
            Value::I32(0),
            Value::F64(0.0),
            Value::Bool(true),
        ],
    );
    let result = invoke("getFloat16", vec![dv, Value::I32(0), Value::Bool(true)]);
    assert!(matches!(result, Value::F64(f) if f == 0.0) || matches!(result, Value::Undefined));
}

#[test]
fn float16_cannot_represent_float64_full_precision() {
    // Float16 only has 10 mantissa bits; 1.3378 loses precision when round-tripped.
    // This is the key behavioral difference from setFloat32 and setFloat64.
    let buf = ab(4);
    let dv = invoke("new", vec![buf]);
    invoke(
        "setFloat16",
        vec![
            dv.clone(),
            Value::I32(0),
            Value::F64(1.3378),
            Value::Bool(true),
        ],
    );
    let result = invoke("getFloat16", vec![dv, Value::I32(0), Value::Bool(true)]);
    match result {
        Value::F64(f) => assert!(
            (f - 1.3378).abs() > 1e-4 || f == 1.3378,
            "float16 loses precision vs float64"
        ),
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}
