//! Tests for `emitter::ops` — the spec-compliant replacements for DYN_* opcodes.
//!
//! Each function emits through the ops.rs sequence, runs the VM with the
//! wasm:js-* host functions registered, and asserts the result.

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, VM, Value};

/// Unique names for test-argument globals, so reused VMs never collide.
static TEST_GLOBAL_SEQ: AtomicUsize = AtomicUsize::new(0);

thread_local! {
    /// Globals queued by `push` for values with no spec const emitter; `run`
    /// drains them into the VM it creates before running the chunk. Each test
    /// runs on its own thread, and `push` only executes inside `run`'s emit
    /// closure, so queue and drain are sequential per test.
    static PENDING_GLOBALS: std::cell::RefCell<Vec<(String, Value)>> =
        const { std::cell::RefCell::new(Vec::new()) };
}

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<test>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    PENDING_GLOBALS.with(|p| {
        for (name, value) in p.borrow_mut().drain(..) {
            vm.globals.insert(name, value);
        }
    });
    vm.run(vec![chunk]).expect("VM run failed")
}

fn push(c: &mut Chunk, v: Value) {
    match v {
        Value::I32(n) => c.emit_i32_const(n, 0),
        Value::I64(n) => c.emit_i64_const(n, 0),
        Value::F32(f) => c.emit_f32_const(f, 0),
        Value::F64(f) => c.emit_f64_const(f, 0),
        Value::Bool(b) => c.emit_bool_const(b, 0),
        Value::String(s) => c.emit_string_const(&s, 0),
        Value::Null => c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let name = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, Ordering::Relaxed)
            );
            let ci = c.intern_string_constant(&name);
            c.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            PENDING_GLOBALS.with(|p| p.borrow_mut().push((name, other)));
        }
    }
}

// ── emit_dyn_to_bool ─────────────────────────────────────────────────

#[test]
fn to_bool_null_is_false() {
    let r = run(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_undefined_is_false() {
    let r = run(|c| {
        push(c, Value::Undefined);
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_zero_is_false() {
    let r = run(|c| {
        push(c, Value::F64(0.0));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_nan_is_false() {
    let r = run(|c| {
        push(c, Value::F64(f64::NAN));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_empty_string_is_false() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("")));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_nonempty_string_is_true() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("x")));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn to_bool_nonzero_number_is_true() {
    let r = run(|c| {
        push(c, Value::F64(1.0));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn to_bool_false_bool_is_false() {
    let r = run(|c| {
        push(c, Value::Bool(false));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_true_bool_is_true() {
    let r = run(|c| {
        push(c, Value::Bool(true));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn to_bool_bigint_zero_is_false() {
    let r = run(|c| {
        push(c, Value::bigint_i64(0));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn to_bool_bigint_nonzero_is_true() {
    let r = run(|c| {
        push(c, Value::bigint_i64(-1));
        ops::emit_dyn_to_bool(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ── emit_dyn_not ─────────────────────────────────────────────────────

#[test]
fn not_true_gives_false() {
    let r = run(|c| {
        push(c, Value::Bool(true));
        ops::emit_dyn_not(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn not_false_gives_true() {
    let r = run(|c| {
        push(c, Value::Bool(false));
        ops::emit_dyn_not(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn not_null_gives_true() {
    let r = run(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        ops::emit_dyn_not(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn not_nonzero_gives_false() {
    let r = run(|c| {
        push(c, Value::F64(42.0));
        ops::emit_dyn_not(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

// ── emit_dyn_eq ──────────────────────────────────────────────────────

#[test]
fn eq_same_numbers() {
    let r = run(|c| {
        push(c, Value::F64(3.0));
        push(c, Value::F64(3.0));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn eq_different_numbers() {
    let r = run(|c| {
        push(c, Value::F64(1.0));
        push(c, Value::F64(2.0));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn eq_nan_is_not_equal_to_itself() {
    let r = run(|c| {
        push(c, Value::F64(f64::NAN));
        push(c, Value::F64(f64::NAN));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn eq_same_strings() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("hello")));
        push(c, Value::String(Arc::from("hello")));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn eq_different_strings() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("a")));
        push(c, Value::String(Arc::from("b")));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn eq_both_null() {
    let r = run(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn eq_null_and_undefined_are_equal() {
    let r = run(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        push(c, Value::Undefined);
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn eq_same_bigints() {
    let r = run(|c| {
        push(c, Value::bigint_i64(42));
        push(c, Value::bigint_i64(42));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn eq_different_bigints() {
    let r = run(|c| {
        push(c, Value::bigint_i64(1));
        push(c, Value::bigint_i64(2));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn eq_true_booleans() {
    let r = run(|c| {
        push(c, Value::Bool(true));
        push(c, Value::Bool(true));
        ops::emit_dyn_eq(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ── emit_dyn_ne ──────────────────────────────────────────────────────

#[test]
fn ne_same_returns_false() {
    let r = run(|c| {
        push(c, Value::F64(5.0));
        push(c, Value::F64(5.0));
        ops::emit_dyn_ne(c, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn ne_different_returns_true() {
    let r = run(|c| {
        push(c, Value::F64(1.0));
        push(c, Value::F64(2.0));
        ops::emit_dyn_ne(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ── emit_dyn_lt / gt / le / ge ───────────────────────────────────────

#[test]
fn lt_numbers() {
    let r = run(|c| {
        push(c, Value::F64(1.0));
        push(c, Value::F64(2.0));
        ops::emit_dyn_lt(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
    let r2 = run(|c| {
        push(c, Value::F64(2.0));
        push(c, Value::F64(1.0));
        ops::emit_dyn_lt(c, 0);
    });
    assert_eq!(r2.as_i32(), 0);
}

#[test]
fn gt_numbers() {
    let r = run(|c| {
        push(c, Value::F64(5.0));
        push(c, Value::F64(2.0));
        ops::emit_dyn_gt(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn le_equal_numbers() {
    let r = run(|c| {
        push(c, Value::F64(3.0));
        push(c, Value::F64(3.0));
        ops::emit_dyn_le(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn ge_greater_number() {
    let r = run(|c| {
        push(c, Value::F64(10.0));
        push(c, Value::F64(9.0));
        ops::emit_dyn_ge(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn lt_strings_lexicographic() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("apple")));
        push(c, Value::String(Arc::from("banana")));
        ops::emit_dyn_lt(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn lt_bigints() {
    let r = run(|c| {
        push(c, Value::bigint_i64(1));
        push(c, Value::bigint_i64(2));
        ops::emit_dyn_lt(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
    let r2 = run(|c| {
        push(c, Value::bigint_i64(5));
        push(c, Value::bigint_i64(2));
        ops::emit_dyn_lt(c, 0);
    });
    assert_eq!(r2.as_i32(), 0);
}

#[test]
fn ge_bigints() {
    let r = run(|c| {
        push(c, Value::bigint_i64(10));
        push(c, Value::bigint_i64(10));
        ops::emit_dyn_ge(c, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ── emit_dyn_add ─────────────────────────────────────────────────────

#[test]
fn add_numbers() {
    let r = run(|c| {
        push(c, Value::F64(3.0));
        push(c, Value::F64(4.0));
        ops::emit_dyn_add(c, 0);
    });
    assert_eq!(r.as_f64(), 7.0);
}

#[test]
fn add_string_concat() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("foo")));
        push(c, Value::String(Arc::from("bar")));
        ops::emit_dyn_add(c, 0);
    });
    assert_eq!(format!("{}", r), "foobar");
}

#[test]
fn add_string_and_number_coerces() {
    let r = run(|c| {
        push(c, Value::String(Arc::from("n=")));
        push(c, Value::F64(42.0));
        ops::emit_dyn_add(c, 0);
    });
    assert_eq!(format!("{}", r), "n=42");
}

#[test]
fn add_bigints() {
    let r = run(|c| {
        push(c, Value::bigint_i64(10));
        push(c, Value::bigint_i64(32));
        ops::emit_dyn_add(c, 0);
    });
    assert_eq!(r.as_i64(), 42);
}

// ── emit_dyn_neg ─────────────────────────────────────────────────────

#[test]
fn neg_number() {
    let r = run(|c| {
        push(c, Value::F64(5.0));
        ops::emit_dyn_neg(c, 0);
    });
    assert_eq!(r.as_f64(), -5.0);
}

#[test]
fn neg_bigint() {
    let r = run(|c| {
        push(c, Value::bigint_i64(7));
        ops::emit_dyn_neg(c, 0);
    });
    assert_eq!(r.as_i64(), -7);
}

#[test]
fn neg_negative_number() {
    let r = run(|c| {
        push(c, Value::F64(-3.0));
        ops::emit_dyn_neg(c, 0);
    });
    assert_eq!(r.as_f64(), 3.0);
}
