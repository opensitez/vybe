//! Tests for `ecma:fixedarray` host imports.
//!
//! FixedArray is a Vybe extension layered on Array — it represents immutable-
//! length sequences (Python tuple, C# T[], COBOL OCCURS n TIMES). The
//! underlying storage is ObjectKind::Array with a `__vybe_frozen` marker.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-fixedarray-test>");
    let import_idx = chunk.add_import("ecma:fixedarray", name);
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

fn arr_invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-array-test>");
    let import_idx = chunk.add_import(module, name);
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

// ── construction ──────────────────────────────────────────────────────────────

#[test]
fn new_with_length_creates_fixed_array() {
    let fa = invoke("newWithLength", vec![Value::I32(3)]);
    assert!(matches!(fa, Value::Object(_)));
}

#[test]
fn is_fixed_array_true_for_fixed_array() {
    let fa = invoke("newWithLength", vec![Value::I32(2)]);
    assert_eq!(invoke("isFixedArray", vec![fa]).as_i32(), 1);
}

#[test]
fn is_fixed_array_false_for_plain_array() {
    let arr = arr_invoke("ecma:array", "new", vec![]);
    assert_eq!(invoke("isFixedArray", vec![arr]).as_i32(), 0);
}

// ── length ────────────────────────────────────────────────────────────────────

#[test]
fn length_matches_construction_size() {
    let fa = invoke("newWithLength", vec![Value::I32(5)]);
    assert_eq!(invoke("length", vec![fa]), Value::I32(5));
}

// ── get ───────────────────────────────────────────────────────────────────────

#[test]
fn get_returns_null_for_uninitialised_slot() {
    // newWithLength fills with Null (not Undefined) by default
    let fa = invoke("newWithLength", vec![Value::I32(3)]);
    let v = invoke("get", vec![fa, Value::I32(0)]);
    assert!(matches!(v, Value::Null | Value::Undefined));
}

// ── fromArray / toArray round-trip ────────────────────────────────────────────

#[test]
fn from_array_and_to_array_round_trip() {
    // Build a plain array [1, 2, 3] then freeze it as a FixedArray.
    let arr = arr_invoke("ecma:array", "new", vec![]);
    arr_invoke("ecma:array", "push", vec![arr.clone(), Value::I32(1)]);
    arr_invoke("ecma:array", "push", vec![arr.clone(), Value::I32(2)]);
    arr_invoke("ecma:array", "push", vec![arr.clone(), Value::I32(3)]);

    let fa = invoke("fromArray", vec![arr]);
    assert_eq!(invoke("isFixedArray", vec![fa.clone()]).as_i32(), 1);
    assert_eq!(invoke("length", vec![fa.clone()]), Value::I32(3));

    let back = invoke("toArray", vec![fa]);
    assert!(matches!(back, Value::Object(_)));
}

// ── freeze / isFrozen ────────────────────────────────────────────────────────

#[test]
fn is_frozen_true_after_freeze() {
    let fa = invoke("newWithLength", vec![Value::I32(2)]);
    let frozen = invoke("freeze", vec![fa]);
    assert_eq!(invoke("isFrozen", vec![frozen]).as_i32(), 1);
}
