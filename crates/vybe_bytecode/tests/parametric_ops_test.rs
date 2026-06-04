//! Tests for WASM parametric instructions (§5.3):
//! unreachable (0x00), nop (0x01), drop (0x1A), select (0x1B), select_t (0x1C).

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).unwrap_err().to_string()
}

fn push_i32(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

// ── unreachable ──────────────────────────────────────────────────────────

#[test]
fn unreachable_traps() {
    let e = run_err(|c| {
        c.emit_op(Op::UNREACHABLE, 0);
    });
    assert!(e.contains("unreachable"));
}

// ── nop ──────────────────────────────────────────────────────────────────

#[test]
fn nop_is_transparent() {
    let r = run(|c| {
        push_i32(c, 42);
        c.emit_op(Op::NOP, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn nop_sequence_is_transparent() {
    let r = run(|c| {
        push_i32(c, 7);
        c.emit_op(Op::NOP, 0);
        c.emit_op(Op::NOP, 0);
        c.emit_op(Op::NOP, 0);
    });
    assert_eq!(r.as_i32(), 7);
}

// ── drop ─────────────────────────────────────────────────────────────────

#[test]
fn drop_removes_top() {
    let r = run(|c| {
        push_i32(c, 99);
        push_i32(c, 42);
        c.emit_op(Op::DROP, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn drop_leaves_deeper_value() {
    let r = run(|c| {
        push_i32(c, 1);
        push_i32(c, 2);
        push_i32(c, 3);
        c.emit_op(Op::DROP, 0);
        c.emit_op(Op::DROP, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ── select ────────────────────────────────────────────────────────────────

#[test]
fn select_picks_first_when_nonzero() {
    // select: [val1, val2, condition] — picks val1 if condition != 0
    let r = run(|c| {
        push_i32(c, 10);
        push_i32(c, 20);
        push_i32(c, 1);
        c.emit_op(Op::SELECT, 0);
    });
    assert_eq!(r.as_i32(), 10);
}

#[test]
fn select_picks_second_when_zero() {
    let r = run(|c| {
        push_i32(c, 10);
        push_i32(c, 20);
        push_i32(c, 0);
        c.emit_op(Op::SELECT, 0);
    });
    assert_eq!(r.as_i32(), 20);
}

#[test]
fn select_with_negative_condition_picks_first() {
    let r = run(|c| {
        push_i32(c, 10);
        push_i32(c, 20);
        push_i32(c, -1);
        c.emit_op(Op::SELECT, 0);
    });
    assert_eq!(r.as_i32(), 10);
}

// ── select_t ─────────────────────────────────────────────────────────────

#[test]
fn select_t_picks_first_when_nonzero() {
    let r = run(|c| {
        push_i32(c, 10);
        push_i32(c, 20);
        push_i32(c, 1);
        c.emit_op(Op::SELECT_T, 0);
    });
    assert_eq!(r.as_i32(), 10);
}

#[test]
fn select_t_picks_second_when_zero() {
    let r = run(|c| {
        push_i32(c, 10);
        push_i32(c, 20);
        push_i32(c, 0);
        c.emit_op(Op::SELECT_T, 0);
    });
    assert_eq!(r.as_i32(), 20);
}
