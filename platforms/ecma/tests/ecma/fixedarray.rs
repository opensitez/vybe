//! Tests for `ecma:fixedarray` host imports.
//!
//! FixedArray is a Vybe extension layered on Array — it represents immutable-
//! length sequences (Python tuple, C# T[], COBOL OCCURS n TIMES). The
//! underlying storage is ObjectKind::Array with a `__vybe_frozen` marker.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.globals.insert(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-fixedarray-test>");
    let import_idx = chunk.add_import("ecma:fixedarray", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn arr_invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-array-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
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
