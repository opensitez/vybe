//! Tests for the threads proposal (0xFE prefix): atomic memory operations.
//! Covers: atomic.fence, i32.atomic.load/store, i32.atomic.rmw.*,
//!         i64.atomic.load/store, i64.atomic.rmw.*, memory.atomic.wait32/notify.
//! All operations run single-threaded; wait32 returns 1 (not-equal), notify returns 0.

use std::thread;
use std::time::Duration;
use vybe_bytecode::shared_memory::SharedMemory;
use vybe_bytecode::value::Value;
use vybe_bytecode::wasm;
use vybe_bytecode::{Chunk, Op, VM};

fn write_leb_u32(out: &mut Vec<u8>, mut value: u32) {
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

fn emit_leb_u64(out: &mut Chunk, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.emit(byte, 0);
        if value == 0 {
            break;
        }
    }
}

fn push_section(out: &mut Vec<u8>, id: u8, payload: &[u8]) {
    out.push(id);
    write_leb_u32(out, payload.len() as u32);
    out.extend_from_slice(payload);
}

fn standard_memory64_threads_module(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    push_section(
        &mut out,
        5,
        &[
            0x01, // one memory
            0x07, // memory64 shared memory with max
            0x01, // min 1 page
            0x01, // max 1 page
        ],
    );

    let mut body = Vec::new();
    body.push(0x00); // local decl count
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn standard_threads_module(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    push_section(
        &mut out,
        5,
        &[
            0x01, // one memory
            0x03, // shared memory with max
            0x01, // min 1 page
            0x01, // max 1 page
        ],
    );

    let mut body = Vec::new();
    body.push(0x00); // local decl count
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn run_with_memory(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0); // 1 page
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    vm.run(vec![c]).expect("run failed")
}

fn run_with_small_memory_err(bytes: usize, emit: impl FnOnce(&mut Chunk)) -> String {
    let mut vm = VM::new();
    vm.memory.resize(bytes, 0);
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    vm.run(vec![c]).unwrap_err().to_string()
}

fn push_i32(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_i64(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

fn emit_atomic(c: &mut Chunk, op: Op) {
    c.emit_op(op, 0);
    c.emit(0, 0); // memarg align
    c.emit(0, 0); // memarg offset
}

fn emit_atomic64(c: &mut Chunk, op: Op, align: u32, offset: u64) {
    c.emit_op(op, 0);
    c.emit_leb_u32(align | 0x80, 0); // memory64 marker in bytecode memarg
    emit_leb_u64(c, offset);
}

fn emit_atomic_fence(c: &mut Chunk) {
    c.emit_op(Op::ATOMIC_FENCE, 0);
    c.emit(0, 0);
}

// ── standard WASM threads opcode decoding ────────────────────────────────

#[test]
fn standard_i32_atomic_load8_u_must_not_decode_as_noop() {
    let bytes = standard_threads_module(&[
        0x41, 0x00, // i32.const 0
        0xFE, 0x12, // i32.atomic.load8_u
        0x00, 0x00, // memarg align=0 offset=0
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("i32.atomic.load8_u should decode");
    assert!(
        chunks[1]
            .code
            .windows(4)
            .any(|w| w == [0x00, 0xFE, 0x00, 0x12])
    );
}

#[test]
fn standard_i64_atomic_rmw_and_must_not_decode_as_noop() {
    let bytes = standard_threads_module(&[
        0x41, 0x00, // i32.const 0
        0x42, 0x01, // i64.const 1
        0xFE, 0x2D, // i64.atomic.rmw.and
        0x03, 0x00, // memarg align=3 offset=0
        0xA7, // i32.wrap_i64 so the function body has i32 result shape
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("i64.atomic.rmw.and should decode");
    assert!(
        chunks[1]
            .code
            .windows(4)
            .any(|w| w == [0x00, 0xFE, 0x00, 0x2D])
    );
}

#[test]
fn all_standard_atomic_opcodes_decode_to_bytecode() {
    let cases: &[(&str, u8, &[u8], &[u8])] = &[
        ("i32.atomic.load8_u", 0x12, &[0x41, 0x00], &[]),
        ("i32.atomic.load16_u", 0x13, &[0x41, 0x00], &[]),
        ("i64.atomic.load8_u", 0x14, &[0x41, 0x00], &[0xA7]),
        ("i64.atomic.load16_u", 0x15, &[0x41, 0x00], &[0xA7]),
        ("i64.atomic.load32_u", 0x16, &[0x41, 0x00], &[0xA7]),
        (
            "i32.atomic.store8",
            0x19,
            &[0x41, 0x00, 0x41, 0x01],
            &[0x41, 0x00],
        ),
        (
            "i32.atomic.store16",
            0x1A,
            &[0x41, 0x00, 0x41, 0x01],
            &[0x41, 0x00],
        ),
        (
            "i64.atomic.store8",
            0x1B,
            &[0x41, 0x00, 0x42, 0x01],
            &[0x41, 0x00],
        ),
        (
            "i64.atomic.store16",
            0x1C,
            &[0x41, 0x00, 0x42, 0x01],
            &[0x41, 0x00],
        ),
        (
            "i64.atomic.store32",
            0x1D,
            &[0x41, 0x00, 0x42, 0x01],
            &[0x41, 0x00],
        ),
        (
            "i64.atomic.rmw.and",
            0x2D,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw.or",
            0x34,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw.xor",
            0x3B,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw.xchg",
            0x42,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i32.atomic.rmw8.add_u",
            0x20,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i32.atomic.rmw16.add_u",
            0x21,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i64.atomic.rmw8.add_u",
            0x22,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.add_u",
            0x23,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.add_u",
            0x24,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i32.atomic.rmw8.sub_u",
            0x27,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i32.atomic.rmw16.sub_u",
            0x28,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i64.atomic.rmw8.sub_u",
            0x29,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.sub_u",
            0x2A,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.sub_u",
            0x2B,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i32.atomic.rmw8.and_u",
            0x2E,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i32.atomic.rmw16.and_u",
            0x2F,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i64.atomic.rmw8.and_u",
            0x30,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.and_u",
            0x31,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.and_u",
            0x32,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        ("i32.atomic.rmw8.or_u", 0x35, &[0x41, 0x00, 0x41, 0x01], &[]),
        (
            "i32.atomic.rmw16.or_u",
            0x36,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i64.atomic.rmw8.or_u",
            0x37,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.or_u",
            0x38,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.or_u",
            0x39,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i32.atomic.rmw8.xor_u",
            0x3C,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i32.atomic.rmw16.xor_u",
            0x3D,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i64.atomic.rmw8.xor_u",
            0x3E,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.xor_u",
            0x3F,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.xor_u",
            0x40,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i32.atomic.rmw8.xchg_u",
            0x43,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i32.atomic.rmw16.xchg_u",
            0x44,
            &[0x41, 0x00, 0x41, 0x01],
            &[],
        ),
        (
            "i64.atomic.rmw8.xchg_u",
            0x45,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.xchg_u",
            0x46,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.xchg_u",
            0x47,
            &[0x41, 0x00, 0x42, 0x01],
            &[0xA7],
        ),
        (
            "i32.atomic.rmw8.cmpxchg_u",
            0x4A,
            &[0x41, 0x00, 0x41, 0x01, 0x41, 0x02],
            &[],
        ),
        (
            "i32.atomic.rmw16.cmpxchg_u",
            0x4B,
            &[0x41, 0x00, 0x41, 0x01, 0x41, 0x02],
            &[],
        ),
        (
            "i64.atomic.rmw8.cmpxchg_u",
            0x4C,
            &[0x41, 0x00, 0x42, 0x01, 0x42, 0x02],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw16.cmpxchg_u",
            0x4D,
            &[0x41, 0x00, 0x42, 0x01, 0x42, 0x02],
            &[0xA7],
        ),
        (
            "i64.atomic.rmw32.cmpxchg_u",
            0x4E,
            &[0x41, 0x00, 0x42, 0x01, 0x42, 0x02],
            &[0xA7],
        ),
    ];

    for (name, subopcode, operands, result_fixup) in cases {
        let mut body = Vec::new();
        body.extend_from_slice(operands);
        body.extend_from_slice(&[0xFE, *subopcode, 0x00, 0x00]);
        body.extend_from_slice(result_fixup);
        let bytes = standard_threads_module(&body);
        let chunks = wasm::read_wasm(&bytes).unwrap_or_else(|err| {
            panic!("{name} should decode to bytecode instead of being skipped: {err}")
        });
        assert!(
            chunks[1]
                .code
                .windows(4)
                .any(|w| w == [0x00, 0xFE, 0x00, *subopcode]),
            "{name} must be present in bytecode"
        );
    }
}

#[test]
fn standard_memory64_atomic_opcodes_decode_with_i64_address_shape() {
    let bytes = standard_memory64_threads_module(&[
        0x42, 0x00, // i64.const 0
        0xFE, 0x10, // i32.atomic.load
        0x02, 0x00, // memarg align=2 offset=0
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("memory64 i32.atomic.load should decode");
    assert!(
        chunks[1]
            .code
            .windows(4)
            .any(|w| w == [0x00, 0xFE, 0x00, 0x10]),
        "standard memory64 atomic opcode must remain the 0xFE i32.atomic.load opcode"
    );
}

#[test]
fn memory64_atomic_store_and_load_use_i64_address_and_u64_offset() {
    let mut vm = VM::new();
    vm.memory.resize(64, 0);
    let mut c = Chunk::new("<memory64-atomic>");
    push_i64(&mut c, 0);
    push_i32(&mut c, 0x1122_3344);
    emit_atomic64(&mut c, Op::I32_ATOMIC_STORE, 2, 8);
    push_i64(&mut c, 0);
    emit_atomic64(&mut c, Op::I32_ATOMIC_LOAD, 2, 8);
    c.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![c]).expect("memory64 atomic run failed");
    assert_eq!(result.as_i32(), 0x1122_3344);
}

// ── atomic.fence ─────────────────────────────────────────────────────────

#[test]
fn atomic_fence_is_noop() {
    // fence emits nothing observable — just verify it doesn't crash
    let r = run_with_memory(|c| {
        emit_atomic_fence(c);
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
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0); // address
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn i32_atomic_store_overwrites() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 2);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 2);
}

// ── i32.atomic.rmw.add ────────────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_add_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 10);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 5);
        emit_atomic(c, Op::I32_ATOMIC_RMW_ADD);
        // returns old value (10)
    });
    assert_eq!(r.as_i32(), 10);
}

#[test]
fn i32_atomic_rmw_add_updates_memory() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 10);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 5);
        emit_atomic(c, Op::I32_ATOMIC_RMW_ADD);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD); // should be 15
    });
    assert_eq!(r.as_i32(), 15);
}

// ── i32.atomic.rmw.sub ────────────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_sub() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 20);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 8);
        emit_atomic(c, Op::I32_ATOMIC_RMW_SUB);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 12);
}

// ── i32.atomic.rmw.and/or/xor ─────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_and() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1100);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 0b1010);
        emit_atomic(c, Op::I32_ATOMIC_RMW_AND);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 0b1000);
}

#[test]
fn i32_atomic_rmw_and_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1100);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 0b1010);
        emit_atomic(c, Op::I32_ATOMIC_RMW_AND);
    });
    assert_eq!(r.as_i32(), 0b1100);
}

#[test]
fn i32_atomic_rmw_or() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1100);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 0b0011);
        emit_atomic(c, Op::I32_ATOMIC_RMW_OR);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 0b1111);
}

#[test]
fn i32_atomic_rmw_or_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1100);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 0b0011);
        emit_atomic(c, Op::I32_ATOMIC_RMW_OR);
    });
    assert_eq!(r.as_i32(), 0b1100);
}

#[test]
fn i32_atomic_rmw_xor() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1111);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 0b1010);
        emit_atomic(c, Op::I32_ATOMIC_RMW_XOR);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 0b0101);
}

#[test]
fn i32_atomic_rmw_xor_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0b1111);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 0b1010);
        emit_atomic(c, Op::I32_ATOMIC_RMW_XOR);
    });
    assert_eq!(r.as_i32(), 0b1111);
}

// ── i32.atomic.rmw.xchg ───────────────────────────────────────────────────

#[test]
fn i32_atomic_rmw_xchg_returns_old() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 7);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 99);
        emit_atomic(c, Op::I32_ATOMIC_RMW_XCHG);
    });
    assert_eq!(r.as_i32(), 7);
}

#[test]
fn i32_atomic_rmw_xchg_updates_memory() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 7);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 99);
        emit_atomic(c, Op::I32_ATOMIC_RMW_XCHG);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 99);
}

// ── i32.atomic.rmw.cmpxchg ────────────────────────────────────────────────

#[test]
fn i32_atomic_cmpxchg_succeeds_when_expected_matches() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 42);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 42);
        push_i32(c, 100);
        emit_atomic(c, Op::I32_ATOMIC_RMW_CMPXCHG);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 100); // replaced
}

#[test]
fn i32_atomic_cmpxchg_fails_when_expected_mismatches() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 42);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 99);
        push_i32(c, 100);
        emit_atomic(c, Op::I32_ATOMIC_RMW_CMPXCHG);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i32(), 42); // unchanged
}

#[test]
fn i32_atomic_cmpxchg_returns_old_on_success() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 42);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 42);
        push_i32(c, 100);
        emit_atomic(c, Op::I32_ATOMIC_RMW_CMPXCHG);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn i32_atomic_cmpxchg_returns_old_on_failure() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 42);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 99);
        push_i32(c, 100);
        emit_atomic(c, Op::I32_ATOMIC_RMW_CMPXCHG);
    });
    assert_eq!(r.as_i32(), 42);
}

// ── i64.atomic.load / i64.atomic.store ───────────────────────────────────

#[test]
fn i64_atomic_store_and_load() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, i64::MAX);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), i64::MAX);
}

// ── i64.atomic.rmw.add / sub ──────────────────────────────────────────────

#[test]
fn i64_atomic_rmw_add() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 100);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 42);
        emit_atomic(c, Op::I64_ATOMIC_RMW_ADD);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), 142);
}

#[test]
fn i64_atomic_rmw_add_returns_old_and_wraps() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, i64::MAX);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_RMW_ADD);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), i64::MIN);
}

#[test]
fn i64_atomic_rmw_sub() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 200);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 58);
        emit_atomic(c, Op::I64_ATOMIC_RMW_SUB);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), 142);
}

#[test]
fn i64_atomic_rmw_sub_returns_old_and_wraps() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, i64::MIN);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_RMW_SUB);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), i64::MAX);
}

// ── i64.atomic.rmw.cmpxchg ────────────────────────────────────────────────

#[test]
fn i64_atomic_cmpxchg_succeeds() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 7);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 7);
        push_i64(c, 99);
        emit_atomic(c, Op::I64_ATOMIC_RMW_CMPXCHG);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), 99);
}

#[test]
fn i64_atomic_cmpxchg_returns_old_on_success() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 7);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 7);
        push_i64(c, 99);
        emit_atomic(c, Op::I64_ATOMIC_RMW_CMPXCHG);
    });
    assert_eq!(r.as_i64(), 7);
}

#[test]
fn i64_atomic_cmpxchg_fails_when_expected_mismatches() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 7);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 8);
        push_i64(c, 99);
        emit_atomic(c, Op::I64_ATOMIC_RMW_CMPXCHG);
        c.emit_op(Op::DROP, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert_eq!(r.as_i64(), 7);
}

#[test]
fn i64_atomic_cmpxchg_returns_old_on_failure() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 7);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0);
        push_i64(c, 8);
        push_i64(c, 99);
        emit_atomic(c, Op::I64_ATOMIC_RMW_CMPXCHG);
    });
    assert_eq!(r.as_i64(), 7);
}

// ── memory.atomic.wait32 / notify ─────────────────────────────────────────

#[test]
fn memory_atomic_wait32_returns_not_equal() {
    // In single-threaded VM: wait32 with value != stored value returns 1 ("not-equal")
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0); // address
        push_i32(c, 999); // expected (doesn't match stored 0)
        push_i64(c, -1); // timeout = -1 (infinite)
        emit_atomic(c, Op::MEMORY_ATOMIC_WAIT32);
    });
    // 1 = "not-equal" (value at address didn't match expected)
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn memory_atomic_wait32_zero_timeout_times_out() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i32(c, 123);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
        push_i32(c, 0);
        push_i32(c, 123);
        push_i64(c, 0);
        emit_atomic(c, Op::MEMORY_ATOMIC_WAIT32);
    });
    assert_eq!(r.as_i32(), 2);
}

#[test]
fn memory_atomic_wait64_returns_not_equal() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0); // address
        push_i64(c, 999); // expected (doesn't match stored 0)
        push_i64(c, -1); // timeout = -1 (infinite)
        emit_atomic(c, Op::MEMORY_ATOMIC_WAIT64);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn memory_atomic_wait64_zero_timeout_times_out() {
    let r = run_with_memory(|c| {
        push_i32(c, 0);
        push_i64(c, 123);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
        push_i32(c, 0); // address
        push_i64(c, 123); // expected matches
        push_i64(c, 0); // timeout = 0
        emit_atomic(c, Op::MEMORY_ATOMIC_WAIT64);
    });
    assert_eq!(r.as_i32(), 2);
}

#[test]
fn memory_atomic_notify_returns_zero_waiters() {
    // Single-threaded: notify wakes 0 threads
    let r = run_with_memory(|c| {
        push_i32(c, 0); // address
        push_i32(c, 10); // count
        emit_atomic(c, Op::MEMORY_ATOMIC_NOTIFY);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn shared_memory_wait32_returns_ok_when_notified() {
    let memory = SharedMemory::new(64);
    memory.store_i32(0, 7).unwrap();
    let waiter_memory = memory.clone();
    let waiter = thread::spawn(move || waiter_memory.wait32(0, 7, 5_000_000_000));

    thread::sleep(Duration::from_millis(20));
    assert_eq!(memory.notify(0, 1), 1);
    assert_eq!(waiter.join().unwrap(), 0);
}

#[test]
fn shared_memory_wait64_returns_ok_when_notified() {
    let memory = SharedMemory::new(64);
    memory.store_i64(8, 11).unwrap();
    let waiter_memory = memory.clone();
    let waiter = thread::spawn(move || waiter_memory.wait64(8, 11, 5_000_000_000));

    thread::sleep(Duration::from_millis(20));
    assert_eq!(memory.notify(8, 1), 1);
    assert_eq!(waiter.join().unwrap(), 0);
}

#[test]
fn shared_memory_notify_wakes_at_most_requested_waiters() {
    let memory = SharedMemory::new(64);
    memory.store_i32(0, 13).unwrap();

    let first_memory = memory.clone();
    let first = thread::spawn(move || first_memory.wait32(0, 13, 5_000_000_000));
    let second_memory = memory.clone();
    let second = thread::spawn(move || second_memory.wait32(0, 13, 5_000_000_000));

    thread::sleep(Duration::from_millis(20));
    assert_eq!(memory.notify(0, 1), 1);
    thread::sleep(Duration::from_millis(20));
    assert_eq!(memory.notify(0, 10), 1);

    assert_eq!(first.join().unwrap(), 0);
    assert_eq!(second.join().unwrap(), 0);
}

#[test]
fn memory_atomic_notify_oob_traps() {
    let err = run_with_small_memory_err(3, |c| {
        push_i32(c, 0);
        push_i32(c, 1);
        emit_atomic(c, Op::MEMORY_ATOMIC_NOTIFY);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

// ── atomics traps ────────────────────────────────────────────────────────

#[test]
fn i32_atomic_load_oob_traps() {
    let err = run_with_small_memory_err(3, |c| {
        push_i32(c, 0);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i32_atomic_store_oob_traps() {
    let err = run_with_small_memory_err(3, |c| {
        push_i32(c, 0);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i64_atomic_load_oob_traps() {
    let err = run_with_small_memory_err(7, |c| {
        push_i32(c, 0);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i64_atomic_store_oob_traps() {
    let err = run_with_small_memory_err(7, |c| {
        push_i32(c, 0);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i32_atomic_load_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_LOAD);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i32_atomic_store_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 1);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_STORE);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i64_atomic_load_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 4);
        emit_atomic(c, Op::I64_ATOMIC_LOAD);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i64_atomic_store_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 4);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_STORE);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i32_atomic_rmw_oob_traps() {
    let err = run_with_small_memory_err(3, |c| {
        push_i32(c, 0);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_RMW_ADD);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i64_atomic_rmw_oob_traps() {
    let err = run_with_small_memory_err(7, |c| {
        push_i32(c, 0);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_RMW_ADD);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i64_atomic_rmw_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 4);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_RMW_ADD);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i32_atomic_rmw_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 1);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_RMW_ADD);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i32_atomic_cmpxchg_oob_traps() {
    let err = run_with_small_memory_err(3, |c| {
        push_i32(c, 0);
        push_i32(c, 0);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_RMW_CMPXCHG);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i32_atomic_cmpxchg_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 1);
        push_i32(c, 0);
        push_i32(c, 1);
        emit_atomic(c, Op::I32_ATOMIC_RMW_CMPXCHG);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn i64_atomic_cmpxchg_oob_traps() {
    let err = run_with_small_memory_err(7, |c| {
        push_i32(c, 0);
        push_i64(c, 0);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_RMW_CMPXCHG);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn i64_atomic_cmpxchg_unaligned_traps() {
    let err = run_with_small_memory_err(16, |c| {
        push_i32(c, 4);
        push_i64(c, 0);
        push_i64(c, 1);
        emit_atomic(c, Op::I64_ATOMIC_RMW_CMPXCHG);
    });
    assert!(err.contains("atomic") && err.contains("unaligned"));
}

#[test]
fn memory_atomic_wait32_oob_traps() {
    let err = run_with_small_memory_err(3, |c| {
        push_i32(c, 0);
        push_i32(c, 0);
        push_i64(c, 0);
        emit_atomic(c, Op::MEMORY_ATOMIC_WAIT32);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}

#[test]
fn memory_atomic_wait64_oob_traps() {
    let err = run_with_small_memory_err(7, |c| {
        push_i32(c, 0);
        push_i64(c, 0);
        push_i64(c, 0);
        emit_atomic(c, Op::MEMORY_ATOMIC_WAIT64);
    });
    assert!(err.contains("atomic") && err.contains("out of bounds"));
}
