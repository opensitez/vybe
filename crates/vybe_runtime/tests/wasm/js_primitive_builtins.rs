//! Tests for the js-primitive-builtins WASM proposal (Stage 1).
//! Spec: `proposals/js-primitive-builtins/proposals/js-primitive-builtins/Overview.md`
//!
//! Covers: wasm:js-{number, boolean, undefined, symbol, bigint}

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_runtime::{Chunk, Op, VM, Value};

/// Unique names for test-argument globals, so reused VMs never collide.
static TEST_GLOBAL_SEQ: AtomicUsize = AtomicUsize::new(0);

/// Emit a spec const for `v` when a const emitter exists; otherwise route the
/// value through a VM global (`global.get` of a unique imported name) and
/// return the pending `(name, value)` binding for the VM that runs the chunk.
fn push_arg(chunk: &mut Chunk, v: Value) -> Option<(String, Value)> {
    match v {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let name = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, Ordering::Relaxed)
            );
            let ci = chunk.intern_string_constant(&name);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            return Some((name, other));
        }
    }
    None
}

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut pending = Vec::new();
    for v in args {
        pending.extend(push_arg(&mut chunk, v));
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    for (name, value) in pending {
        vm.globals.insert(name, value);
    }
    vm.run(vec![chunk]).expect("VM run failed")
}

fn call_import_expect_trap(module: &str, name: &str, args: Vec<Value>) {
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut pending = Vec::new();
    for v in args {
        pending.extend(push_arg(&mut chunk, v));
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    for (name, value) in pending {
        vm.globals.insert(name, value);
    }
    assert!(vm.run(vec![chunk]).is_err(), "{module}.{name} should trap");
}

// ── wasm:js-number ────────────────────────────────────────────────────

#[test]
fn js_number_test_recognises_numeric_values() {
    assert_eq!(
        call_import("wasm:js-number", "test", vec![Value::F64(3.14)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "test", vec![Value::I32(0)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "test", vec![Value::I64(99)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "test", vec![Value::Null]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "test", vec![Value::Bool(true)]).as_i32(),
        0
    );
    assert_eq!(
        call_import(
            "wasm:js-number",
            "test",
            vec![Value::String(Arc::from("3"))]
        )
        .as_i32(),
        0
    );
}

#[test]
fn js_number_test_i32_detects_integer_fit() {
    assert_eq!(
        call_import("wasm:js-number", "testI32", vec![Value::I32(42)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "testI32", vec![Value::F64(7.0)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "testI32", vec![Value::F64(7.5)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testI32", vec![Value::F64(2147483648.0)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testI32", vec![Value::Null]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testI32", vec![Value::F64(-0.0)]).as_i32(),
        0
    );
}

#[test]
fn js_number_test_u32_detects_uint_fit() {
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::F64(0.0)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::F64(4294967295.0)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::F64(4294967296.0)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::F64(-1.0)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::F64(-0.0)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::F64(1.5)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "testU32", vec![Value::Null]).as_i32(),
        0
    );
}

#[test]
fn js_number_from_f64_boxes_float() {
    assert_eq!(
        call_import("wasm:js-number", "fromF64", vec![Value::F64(2.5)]).as_f64(),
        2.5
    );
    assert_eq!(
        call_import("wasm:js-number", "fromF64", vec![Value::F64(-0.0)]).as_f64(),
        -0.0
    );
}

#[test]
fn js_number_from_i32_boxes_integer() {
    assert_eq!(
        call_import("wasm:js-number", "fromI32", vec![Value::I32(-7)]).as_i32(),
        -7
    );
    assert_eq!(
        call_import("wasm:js-number", "fromI32", vec![Value::I32(i32::MIN)]).as_i32(),
        i32::MIN
    );
}

#[test]
fn js_number_from_u32_reinterprets_as_unsigned() {
    assert_eq!(
        call_import("wasm:js-number", "fromU32", vec![Value::I32(-1)]).as_f64(),
        4294967295.0
    );
    assert_eq!(
        call_import("wasm:js-number", "fromU32", vec![Value::I32(0)]).as_f64(),
        0.0
    );
}

#[test]
fn js_number_to_f64_unboxes_number() {
    assert_eq!(
        call_import("wasm:js-number", "toF64", vec![Value::F64(1.5)]).as_f64(),
        1.5
    );
    assert_eq!(
        call_import("wasm:js-number", "toF64", vec![Value::I32(42)]).as_f64(),
        42.0
    );
}

#[test]
fn js_number_to_f64_traps_on_non_number() {
    call_import_expect_trap("wasm:js-number", "toF64", vec![Value::Null]);
    call_import_expect_trap("wasm:js-number", "toF64", vec![Value::Bool(true)]);
}

#[test]
fn js_number_to_i32_unboxes_integer() {
    assert_eq!(
        call_import("wasm:js-number", "toI32", vec![Value::I32(42)]).as_i32(),
        42
    );
    assert_eq!(
        call_import("wasm:js-number", "toI32", vec![Value::F64(7.0)]).as_i32(),
        7
    );
    assert_eq!(
        call_import("wasm:js-number", "toI32", vec![Value::F64(-1.0)]).as_i32(),
        -1
    );
}

#[test]
fn js_number_to_i32_traps_on_fractional() {
    call_import_expect_trap("wasm:js-number", "toI32", vec![Value::F64(7.5)]);
}

#[test]
fn js_number_to_i32_traps_on_neg_zero() {
    call_import_expect_trap("wasm:js-number", "toI32", vec![Value::F64(-0.0)]);
}

#[test]
fn js_number_to_i32_traps_on_out_of_range() {
    call_import_expect_trap("wasm:js-number", "toI32", vec![Value::F64(2147483648.0)]);
}

#[test]
fn js_number_to_u32_unboxes_uint() {
    assert_eq!(
        call_import("wasm:js-number", "toU32", vec![Value::F64(4294967295.0)]).as_i32(),
        -1_i32
    );
    assert_eq!(
        call_import("wasm:js-number", "toU32", vec![Value::F64(0.0)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-number", "toU32", vec![Value::F64(1.0)]).as_i32(),
        1
    );
}

#[test]
fn js_number_to_u32_traps_on_invalid() {
    call_import_expect_trap("wasm:js-number", "toU32", vec![Value::F64(-1.0)]);
    call_import_expect_trap("wasm:js-number", "toU32", vec![Value::F64(4294967296.0)]);
    call_import_expect_trap("wasm:js-number", "toU32", vec![Value::F64(1.5)]);
    call_import_expect_trap("wasm:js-number", "toU32", vec![Value::Null]);
}

// ── wasm:js-boolean ───────────────────────────────────────────────────

#[test]
fn js_boolean_test_recognises_bool() {
    assert_eq!(
        call_import("wasm:js-boolean", "test", vec![Value::Bool(true)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-boolean", "test", vec![Value::Bool(false)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-boolean", "test", vec![Value::I32(1)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-boolean", "test", vec![Value::Null]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-boolean", "test", vec![Value::Undefined]).as_i32(),
        0
    );
}

#[test]
fn js_boolean_cast_extracts_i32() {
    assert_eq!(
        call_import("wasm:js-boolean", "cast", vec![Value::Bool(true)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-boolean", "cast", vec![Value::Bool(false)]).as_i32(),
        0
    );
}

#[test]
fn js_boolean_cast_traps_on_non_bool() {
    call_import_expect_trap("wasm:js-boolean", "cast", vec![Value::I32(1)]);
    call_import_expect_trap("wasm:js-boolean", "cast", vec![Value::Null]);
}

// ── wasm:js-undefined ─────────────────────────────────────────────────

#[test]
fn js_undefined_test_recognises_undefined() {
    assert_eq!(
        call_import("wasm:js-undefined", "test", vec![Value::Undefined]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-undefined", "test", vec![Value::Null]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-undefined", "test", vec![Value::Bool(false)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-undefined", "test", vec![Value::I32(0)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-undefined", "test", vec![Value::F64(0.0)]).as_i32(),
        0
    );
}

// ── wasm:js-symbol ────────────────────────────────────────────────────

#[test]
fn js_symbol_test_recognises_symbol() {
    assert_eq!(
        call_import(
            "wasm:js-symbol",
            "test",
            vec![Value::Symbol(Arc::from("foo"))]
        )
        .as_i32(),
        1
    );
    assert_eq!(
        call_import(
            "wasm:js-symbol",
            "test",
            vec![Value::String(Arc::from("foo"))]
        )
        .as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-symbol", "test", vec![Value::Null]).as_i32(),
        0
    );
}

#[test]
fn js_symbol_equals_same_arc_is_true() {
    let arc: Arc<str> = Arc::from("unique");
    let a = Value::Symbol(arc.clone());
    let b = Value::Symbol(arc);
    assert_eq!(
        call_import("wasm:js-symbol", "equals", vec![a, b]).as_i32(),
        1
    );
}

#[test]
fn js_symbol_equals_different_arcs_is_false() {
    let a = Value::Symbol(Arc::from("foo"));
    let b = Value::Symbol(Arc::from("foo"));
    assert_eq!(
        call_import("wasm:js-symbol", "equals", vec![a, b]).as_i32(),
        0
    );
}

#[test]
fn js_symbol_equals_null_null_is_true() {
    assert_eq!(
        call_import("wasm:js-symbol", "equals", vec![Value::Null, Value::Null]).as_i32(),
        1
    );
}

#[test]
fn js_symbol_equals_traps_on_non_symbol_non_null() {
    let sym = Value::Symbol(Arc::from("x"));
    call_import_expect_trap(
        "wasm:js-symbol",
        "equals",
        vec![sym, Value::String(Arc::from("x"))],
    );
}

// ── wasm:js-bigint ────────────────────────────────────────────────────

#[test]
fn js_bigint_test_recognises_bigint() {
    assert_eq!(
        call_import("wasm:js-bigint", "test", vec![Value::bigint_i64(42)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-bigint", "test", vec![Value::bigint_i64(0)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-bigint", "test", vec![Value::bigint_i64(-1)]).as_i32(),
        1
    );
    assert_eq!(
        call_import("wasm:js-bigint", "test", vec![Value::I64(42)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-bigint", "test", vec![Value::I32(42)]).as_i32(),
        0
    );
    assert_eq!(
        call_import("wasm:js-bigint", "test", vec![Value::Null]).as_i32(),
        0
    );
}
