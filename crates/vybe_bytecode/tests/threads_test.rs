//! Tests for the threads proposal (0xFE prefix): atomic memory operations.
//! Covers: atomic.fence, i32.atomic.load/store, i32.atomic.rmw.*,
//!         i64.atomic.load/store, i64.atomic.rmw.*, memory.atomic.wait32/notify.
//! All operations run single-threaded; wait32 returns 1 (not-equal), notify returns 0.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

fn run_with_memory(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0); // 1 page
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    vm.run(vec![c]).expect("run failed")
}

fn push_i32(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_i64(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

// ── atomic.fence ─────────────────────────────────────────────────────────

#[test]
fn atomic_fence_is_noop() {
    // fence emits nothing observable — just verify it doesn't crash
    let r = run_with_memory(|c| {
        c.emit_op(Op::ATOMIC_FENCE, 0);
        push_i32(c, 42);
    });
    assert_eq!(r.as_i32(), 42);
}

// ── i32.atomic.load / i32.atomic.store ───────────────────────────────────

#[test]
fn i32_atomic_store_and_load() {
    let r = run_with_memory(|c| {
        push_i32(c, 0); // address
        push_i32(c, 99); // value
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0); // address
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn i32_atomic_store_overwrites() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 1);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 2);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 2);
}

// ── i32.atomic.rmw.add ────────────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_add_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 10);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 5);
        c.emit_op(Op::I32_ATOMIC_RMW_ADD, 0);
        // returns old value (10)
    });
    assert_eq!(r.as_i32(), 10);
}

#[test]
fn i32_atomic_rmw_add_updates_memory() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 10);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 5);
        c.emit_op(Op::I32_ATOMIC_RMW_ADD, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0); // should be 15
    });
    assert_eq!(r.as_i32(), 15);
}

// ── i32.atomic.rmw.sub ────────────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_sub() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 20);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 8);
        c.emit_op(Op::I32_ATOMIC_RMW_SUB, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 12);
}

// ── i32.atomic.rmw.and/or/xor ─────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_and() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1100);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 0b1010);
        c.emit_op(Op::I32_ATOMIC_RMW_AND, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 0b1000);
}

#[test]
fn i32_atomic_rmw_or() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1100);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 0b0011);
        c.emit_op(Op::I32_ATOMIC_RMW_OR, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 0b1111);
}

#[test]
fn i32_atomic_rmw_xor() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1111);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 0b1010);
        c.emit_op(Op::I32_ATOMIC_RMW_XOR, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 0b0101);
}

// ── i32.atomic.rmw.xchg ───────────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_xchg_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 7);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 99);
        c.emit_op(Op::I32_ATOMIC_RMW_XCHG, 0);
    });
    assert_eq!(r.as_i32(), 7);
}

// ── i32.atomic.rmw.cmpxchg ────────────────────────────────────────────────

#[test]
fn i32_atomic_cmpxchg_succeeds_when_expected_matches() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 42);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 42);
        push_i32(c, 100);
        c.emit_op(Op::I32_ATOMIC_RMW_CMPXCHG, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 100); // replaced
}

#[test]
fn i32_atomic_cmpxchg_fails_when_expected_mismatches() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 42);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i32(c, 99);
        push_i32(c, 100);
        c.emit_op(Op::I32_ATOMIC_RMW_CMPXCHG, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i32(), 42); // unchanged
}

// ── i64.atomic.load / i64.atomic.store ───────────────────────────────────

#[test]
fn i64_atomic_store_and_load() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, i64::MAX);
        c.emit_op(Op::I64_ATOMIC_STORE, 0);
        push_i32(c, 0);
        c.emit_op(Op::I64_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i64(), i64::MAX);
}

// ── i64.atomic.rmw.add / sub ──────────────────────────────────────────────

#[test]
fn i64_atomic_rmw_add() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 100);
        c.emit_op(Op::I64_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i64(c, 42);
        c.emit_op(Op::I64_ATOMIC_RMW_ADD, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I64_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i64(), 142);
}

#[test]
fn i64_atomic_rmw_sub() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 200);
        c.emit_op(Op::I64_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i64(c, 58);
        c.emit_op(Op::I64_ATOMIC_RMW_SUB, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I64_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i64(), 142);
}

// ── i64.atomic.rmw.cmpxchg ────────────────────────────────────────────────

#[test]
fn i64_atomic_cmpxchg_succeeds() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 7);
        c.emit_op(Op::I64_ATOMIC_STORE, 0);
        push_i32(c, 0);
        push_i64(c, 7);
        push_i64(c, 99);
        c.emit_op(Op::I64_ATOMIC_RMW_CMPXCHG, 0);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        c.emit_op(Op::I64_ATOMIC_LOAD, 0);
    });
    assert_eq!(r.as_i64(), 99);
}

// ── memory.atomic.wait32 / notify ─────────────────────────────────────────

#[test]
fn memory_atomic_wait32_returns_not_equal() {
    // In single-threaded VM: wait32 with value != stored value returns 1 ("not-equal")
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0);
        c.emit_op(Op::I32_ATOMIC_STORE, 0);
        push_i32(c, 0); // address
        push_i32(c, 999); // expected (doesn't match stored 0)
        push_i64(c, -1); // timeout = -1 (infinite)
        c.emit_op(Op::MEMORY_ATOMIC_WAIT32, 0);
    });
    // 1 = "not-equal" (value at address didn't match expected)
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn memory_atomic_wait64_returns_not_equal() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 0);
        c.emit_op(Op::I64_ATOMIC_STORE, 0);
        push_i32(c, 0); // address
        push_i64(c, 999); // expected (doesn't match stored 0)
        push_i64(c, -1); // timeout = -1 (infinite)
        c.emit_op(Op::MEMORY_ATOMIC_WAIT64, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn memory_atomic_wait64_zero_timeout_times_out() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 123);
        c.emit_op(Op::I64_ATOMIC_STORE, 0);
        push_i32(c, 0); // address
        push_i64(c, 123); // expected matches
        push_i64(c, 0); // timeout = 0
        c.emit_op(Op::MEMORY_ATOMIC_WAIT64, 0);
    });
    assert_eq!(r.as_i32(), 2);
}

#[test]
fn memory_atomic_notify_returns_zero_waiters() {
    // Single-threaded: notify wakes 0 threads
    let r = run_with_memory(|c| {
        push_i32(c, 0); // address
        push_i32(c, 10); // count
        c.emit_op(Op::MEMORY_ATOMIC_NOTIFY, 0);
    });
    assert_eq!(r.as_i32(), 0);
}
