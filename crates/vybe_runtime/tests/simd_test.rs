//! Tests for the SIMD proposal (0xFD prefix).
//! Covers: v128.const, v128.load/store, v128 bitwise (and/or/xor/not/andnot/any_true/bitselect),
//!         i8x16 (splat/add/sub/eq/extract_lane_s/u/replace_lane/shuffle/swizzle),
//!         i16x8 (splat/add/sub/mul/extract_lane_s/u/replace_lane),
//!         i32x4 (splat/add/sub/mul/eq/gt_s/lt_s/shl/shr_s/shr_u/extract_lane/replace_lane),
//!         f32x4 (splat/add/sub/mul/div/extract_lane/replace_lane),
//!         f64x2 (splat/add/sub/mul/div/min/max/eq/lt/le/sqrt/abs/neg/extract_lane/replace_lane).

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn push_i32(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_i64(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_f64(c: &mut Chunk, v: f64) {
    let k = c.add_constant(Value::F64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

fn emit_leb_u64(c: &mut Chunk, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        c.emit(byte, 0);
        if value == 0 {
            break;
        }
    }
}

fn emit_simd_memarg64(c: &mut Chunk, align: u32, offset: u64) {
    c.emit_leb_u32(align | 0x80 | 0x100, 0);
    emit_leb_u64(c, offset);
}

fn emit_v128_const(c: &mut Chunk, bytes: [u8; 16]) {
    c.emit_op(Op::V128_CONST, 0);
    for &b in &bytes {
        c.emit(b, 0);
    }
}

fn as_v128(v: Value) -> [u8; 16] {
    match v {
        Value::V128(b) => b,
        _ => panic!("expected V128, got {:?}", v) }
}

fn i32_lanes(b: &[u8; 16]) -> [i32; 4] {
    [
        i32::from_le_bytes([b[0], b[1], b[2], b[3]]),
        i32::from_le_bytes([b[4], b[5], b[6], b[7]]),
        i32::from_le_bytes([b[8], b[9], b[10], b[11]]),
        i32::from_le_bytes([b[12], b[13], b[14], b[15]]),
    ]
}

fn f64_lanes(b: &[u8; 16]) -> [f64; 2] {
    [
        f64::from_le_bytes([b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7]]),
        f64::from_le_bytes([b[8], b[9], b[10], b[11], b[12], b[13], b[14], b[15]]),
    ]
}

fn f32_lanes(b: &[u8; 16]) -> [f32; 4] {
    [
        f32::from_le_bytes([b[0], b[1], b[2], b[3]]),
        f32::from_le_bytes([b[4], b[5], b[6], b[7]]),
        f32::from_le_bytes([b[8], b[9], b[10], b[11]]),
        f32::from_le_bytes([b[12], b[13], b[14], b[15]]),
    ]
}

// ── v128.const ────────────────────────────────────────────────────────────

#[test]
fn v128_const_loads_16_bytes() {
    let bytes: [u8; 16] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
    let r = as_v128(run(|c| emit_v128_const(c, bytes)));
    assert_eq!(r, bytes);
}

// ── v128 bitwise ─────────────────────────────────────────────────────────

#[test]
fn v128_and() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0xFF; 16]);
        emit_v128_const(c, [0x0F; 16]);
        c.emit_op(Op::V128_AND, 0);
    }));
    assert!(r.iter().all(|&b| b == 0x0F));
}

#[test]
fn v128_or() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0xF0; 16]);
        emit_v128_const(c, [0x0F; 16]);
        c.emit_op(Op::V128_OR, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF));
}

#[test]
fn v128_xor() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0xFF; 16]);
        emit_v128_const(c, [0xFF; 16]);
        c.emit_op(Op::V128_XOR, 0);
    }));
    assert!(r.iter().all(|&b| b == 0));
}

#[test]
fn v128_not() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0x00; 16]);
        c.emit_op(Op::V128_NOT, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF));
}

#[test]
fn v128_andnot() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0xFF; 16]);
        emit_v128_const(c, [0x0F; 16]);
        c.emit_op(Op::V128_ANDNOT, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xF0));
}

#[test]
fn v128_any_true_nonzero() {
    let r = run(|c| {
        emit_v128_const(c, [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]);
        c.emit_op(Op::V128_ANY_TRUE, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn v128_any_true_all_zero() {
    let r = run(|c| {
        emit_v128_const(c, [0; 16]);
        c.emit_op(Op::V128_ANY_TRUE, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn v128_bitselect() {
    let v1 = [0xFFu8; 16];
    let v2 = [0x00u8; 16];
    let mask = [0xF0u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, v1);
        emit_v128_const(c, v2);
        emit_v128_const(c, mask);
        c.emit_op(Op::V128_BITSELECT, 0);
    }));
    // v1[i] & mask[i] | v2[i] & ~mask[i] = 0xFF & 0xF0 | 0x00 & 0x0F = 0xF0
    assert!(r.iter().all(|&b| b == 0xF0));
    let _ = (v1, v2);
}

// ── v128.load / v128.store ────────────────────────────────────────────────

#[test]
fn v128_store_and_load_roundtrip() {
    let mut vm = VM::new();
    vm.memory.resize(64, 0);
    let bytes: [u8; 16] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
    let mut c = Chunk::new("<script>");
    push_i32(&mut c, 0); // store address
    emit_v128_const(&mut c, bytes);
    c.emit_op(Op::V128_STORE, 0);
    push_i32(&mut c, 0); // load address
    c.emit_op(Op::V128_LOAD, 0);
    c.emit_op(Op::RETURN, 0);
    let r = as_v128(vm.run(vec![c]).expect("run failed"));
    assert_eq!(r, bytes);
}

#[test]
fn memory64_v128_store_and_load_use_i64_address_and_u64_offset() {
    let mut vm = VM::new();
    vm.memory.resize(64, 0);
    let bytes: [u8; 16] = [
        0x10, 0x11, 0x12, 0x13, 0x20, 0x21, 0x22, 0x23, 0x30, 0x31, 0x32, 0x33, 0x40, 0x41, 0x42,
        0x43,
    ];
    let mut c = Chunk::new("<simd-memory64>");
    push_i64(&mut c, 0);
    emit_v128_const(&mut c, bytes);
    c.emit_op(Op::V128_STORE, 0);
    emit_simd_memarg64(&mut c, 4, 8);
    push_i64(&mut c, 0);
    c.emit_op(Op::V128_LOAD, 0);
    emit_simd_memarg64(&mut c, 4, 8);
    c.emit_op(Op::RETURN, 0);

    let result = as_v128(vm.run(vec![c]).expect("memory64 SIMD run failed"));
    assert_eq!(result, bytes);
}

// ── i8x16 ────────────────────────────────────────────────────────────────

#[test]
fn i8x16_splat() {
    let r = as_v128(run(|c| {
        push_i32(c, 7);
        c.emit_op(Op::I8X16_SPLAT, 0);
    }));
    assert!(r.iter().all(|&b| b == 7));
}

#[test]
fn i8x16_add() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [1; 16]);
        emit_v128_const(c, [2; 16]);
        c.emit_op(Op::I8X16_ADD, 0);
    }));
    assert!(r.iter().all(|&b| b == 3));
}

#[test]
fn i8x16_sub() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [5; 16]);
        emit_v128_const(c, [3; 16]);
        c.emit_op(Op::I8X16_SUB, 0);
    }));
    assert!(r.iter().all(|&b| b == 2));
}

#[test]
fn i8x16_eq_true() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [42; 16]);
        emit_v128_const(c, [42; 16]);
        c.emit_op(Op::I8X16_EQ, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF));
}

#[test]
fn i8x16_extract_lane_s() {
    let mut bytes = [0u8; 16];
    bytes[3] = 0xFF; // -1 as i8
    let r = run(|c| {
        emit_v128_const(c, bytes);
        c.emit_op(Op::I8X16_EXTRACT_LANE_S, 0);
        c.emit(3u8, 0); // lane 3
    });
    assert_eq!(r.as_i32(), -1);
}

#[test]
fn i8x16_extract_lane_u() {
    let mut bytes = [0u8; 16];
    bytes[5] = 200;
    let r = run(|c| {
        emit_v128_const(c, bytes);
        c.emit_op(Op::I8X16_EXTRACT_LANE_U, 0);
        c.emit(5u8, 0);
    });
    assert_eq!(r.as_i32(), 200);
}

#[test]
fn i8x16_replace_lane() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0; 16]);
        push_i32(c, 99);
        c.emit_op(Op::I8X16_REPLACE_LANE, 0);
        c.emit(7u8, 0); // lane 7
    }));
    assert_eq!(r[7], 99);
    assert!(
        r.iter()
            .enumerate()
            .filter(|&(i, _)| i != 7)
            .all(|(_, &b)| b == 0)
    );
}

#[test]
fn i8x16_shuffle_identity() {
    let bytes: [u8; 16] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
    let identity: [u8; 16] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
    let r = as_v128(run(|c| {
        emit_v128_const(c, bytes);
        emit_v128_const(c, [0; 16]);
        c.emit_op(Op::I8X16_SHUFFLE, 0);
        for &idx in &identity {
            c.emit(idx, 0);
        }
    }));
    assert_eq!(r, bytes);
}

#[test]
fn i8x16_swizzle() {
    let a: [u8; 16] = [10, 20, 30, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    let indices: [u8; 16] = [2, 0, 1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // pick a[2],a[0],a[1],a[3]
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, indices);
        c.emit_op(Op::I8X16_SWIZZLE, 0);
    }));
    assert_eq!(r[0], 30);
    assert_eq!(r[1], 10);
    assert_eq!(r[2], 20);
    assert_eq!(r[3], 40);
}

// ── i16x8 ────────────────────────────────────────────────────────────────

#[test]
fn i16x8_splat() {
    let r = as_v128(run(|c| {
        push_i32(c, 1000);
        c.emit_op(Op::I16X8_SPLAT, 0);
    }));
    for i in 0..8 {
        let v = i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]);
        assert_eq!(v, 1000);
    }
}

#[test]
fn i16x8_add() {
    let mut a = [0u8; 16];
    let mut b = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&100i16.to_le_bytes());
        b[i * 2..i * 2 + 2].copy_from_slice(&200i16.to_le_bytes());
    }
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I16X8_ADD, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 300);
    }
}

#[test]
fn i16x8_mul() {
    let mut a = [0u8; 16];
    let mut b = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&3i16.to_le_bytes());
        b[i * 2..i * 2 + 2].copy_from_slice(&4i16.to_le_bytes());
    }
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I16X8_MUL, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 12);
    }
}

#[test]
fn i16x8_extract_lane_s() {
    let mut bytes = [0u8; 16];
    bytes[4..6].copy_from_slice(&(-500i16).to_le_bytes());
    let r = run(|c| {
        emit_v128_const(c, bytes);
        c.emit_op(Op::I16X8_EXTRACT_LANE_S, 0);
        c.emit(2u8, 0);
    });
    assert_eq!(r.as_i32(), -500);
}

#[test]
fn i16x8_replace_lane() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0; 16]);
        push_i32(c, 42);
        c.emit_op(Op::I16X8_REPLACE_LANE, 0);
        c.emit(3u8, 0);
    }));
    assert_eq!(i16::from_le_bytes([r[6], r[7]]), 42);
}

// ── i32x4 ────────────────────────────────────────────────────────────────

#[test]
fn i32x4_splat() {
    let r = as_v128(run(|c| {
        push_i32(c, 42);
        c.emit_op(Op::I32X4_SPLAT, 0);
    }));
    assert_eq!(i32_lanes(&r), [42, 42, 42, 42]);
}

#[test]
fn i32x4_add() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&10i32.to_le_bytes());
        }
        b
    };
    let bb: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&32i32.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, bb);
        c.emit_op(Op::I32X4_ADD, 0);
    }));
    assert_eq!(i32_lanes(&r), [42, 42, 42, 42]);
}

#[test]
fn i32x4_sub() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&50i32.to_le_bytes());
        }
        b
    };
    let bb: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&8i32.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, bb);
        c.emit_op(Op::I32X4_SUB, 0);
    }));
    assert_eq!(i32_lanes(&r), [42, 42, 42, 42]);
}

#[test]
fn i32x4_mul() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&6i32.to_le_bytes());
        }
        b
    };
    let bb: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&7i32.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, bb);
        c.emit_op(Op::I32X4_MUL, 0);
    }));
    assert_eq!(i32_lanes(&r), [42, 42, 42, 42]);
}

#[test]
fn i32x4_eq_matching() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&5i32.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_EQ, 0);
    }));
    assert_eq!(i32_lanes(&r), [-1, -1, -1, -1]);
}

#[test]
fn i32x4_shl() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&1i32.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        push_i32(c, 3);
        c.emit_op(Op::I32X4_SHL, 0);
    }));
    assert_eq!(i32_lanes(&r), [8, 8, 8, 8]);
}

#[test]
fn i32x4_shr_s() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&(-8i32).to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        push_i32(c, 1);
        c.emit_op(Op::I32X4_SHR_S, 0);
    }));
    assert_eq!(i32_lanes(&r), [-4, -4, -4, -4]);
}

#[test]
fn i32x4_shr_u() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&8i32.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        push_i32(c, 1);
        c.emit_op(Op::I32X4_SHR_U, 0);
    }));
    assert_eq!(i32_lanes(&r), [4, 4, 4, 4]);
}

#[test]
fn i32x4_extract_lane() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        b[8..12].copy_from_slice(&99i32.to_le_bytes());
        b
    };
    let r = run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_EXTRACT_LANE, 0);
        c.emit(2u8, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn i32x4_replace_lane() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0; 16]);
        push_i32(c, 77);
        c.emit_op(Op::I32X4_REPLACE_LANE, 0);
        c.emit(1u8, 0);
    }));
    assert_eq!(i32_lanes(&r), [0, 77, 0, 0]);
}

// ── f32x4 ────────────────────────────────────────────────────────────────

#[test]
fn f32x4_splat() {
    let r = as_v128(run(|c| {
        push_f64(c, 3.0);
        c.emit_op(Op::F32X4_SPLAT, 0);
    }));
    assert_eq!(f32_lanes(&r), [3.0, 3.0, 3.0, 3.0]);
}

#[test]
fn f32x4_add() {
    let mk = |v: f32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(1.5));
        emit_v128_const(c, mk(2.5));
        c.emit_op(Op::F32X4_ADD, 0);
    }));
    assert_eq!(f32_lanes(&r), [4.0, 4.0, 4.0, 4.0]);
}

#[test]
fn f32x4_mul() {
    let mk = |v: f32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(3.0));
        emit_v128_const(c, mk(4.0));
        c.emit_op(Op::F32X4_MUL, 0);
    }));
    assert_eq!(f32_lanes(&r), [12.0, 12.0, 12.0, 12.0]);
}

#[test]
fn f32x4_extract_lane() {
    let mut bytes = [0u8; 16];
    bytes[4..8].copy_from_slice(&7.0f32.to_le_bytes());
    let r = run(|c| {
        emit_v128_const(c, bytes);
        c.emit_op(Op::F32X4_EXTRACT_LANE, 0);
        c.emit(1u8, 0);
    });
    assert_eq!(r.as_f64() as f32, 7.0);
}

#[test]
fn f32x4_replace_lane() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0; 16]);
        push_f64(c, 5.0);
        c.emit_op(Op::F32X4_REPLACE_LANE, 0);
        c.emit(2u8, 0);
    }));
    assert_eq!(f32_lanes(&r)[2], 5.0);
}

// ── f64x2 ────────────────────────────────────────────────────────────────

#[test]
fn f64x2_splat() {
    let r = as_v128(run(|c| {
        push_f64(c, 3.14);
        c.emit_op(Op::F64X2_SPLAT, 0);
    }));
    let lanes = f64_lanes(&r);
    assert!((lanes[0] - 3.14).abs() < 1e-10);
    assert!((lanes[1] - 3.14).abs() < 1e-10);
}

#[test]
fn f64x2_add() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(1.0));
        emit_v128_const(c, mk(2.0));
        c.emit_op(Op::F64X2_ADD, 0);
    }));
    assert_eq!(f64_lanes(&r), [3.0, 3.0]);
}

#[test]
fn f64x2_sub() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(5.0));
        emit_v128_const(c, mk(3.0));
        c.emit_op(Op::F64X2_SUB, 0);
    }));
    assert_eq!(f64_lanes(&r), [2.0, 2.0]);
}

#[test]
fn f64x2_mul() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(3.0));
        emit_v128_const(c, mk(4.0));
        c.emit_op(Op::F64X2_MUL, 0);
    }));
    assert_eq!(f64_lanes(&r), [12.0, 12.0]);
}

#[test]
fn f64x2_div() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(6.0));
        emit_v128_const(c, mk(2.0));
        c.emit_op(Op::F64X2_DIV, 0);
    }));
    assert_eq!(f64_lanes(&r), [3.0, 3.0]);
}

#[test]
fn f64x2_min_max() {
    let mk = |a: f64, b: f64| {
        let mut buf = [0u8; 16];
        buf[0..8].copy_from_slice(&a.to_le_bytes());
        buf[8..16].copy_from_slice(&b.to_le_bytes());
        buf
    };
    let r_min = as_v128(run(|c| {
        emit_v128_const(c, mk(1.0, 5.0));
        emit_v128_const(c, mk(3.0, 2.0));
        c.emit_op(Op::F64X2_MIN, 0);
    }));
    assert_eq!(f64_lanes(&r_min), [1.0, 2.0]);
    let r_max = as_v128(run(|c| {
        emit_v128_const(c, mk(1.0, 5.0));
        emit_v128_const(c, mk(3.0, 2.0));
        c.emit_op(Op::F64X2_MAX, 0);
    }));
    assert_eq!(f64_lanes(&r_max), [3.0, 5.0]);
}

#[test]
fn f64x2_eq_matching() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(1.0));
        emit_v128_const(c, mk(1.0));
        c.emit_op(Op::F64X2_EQ, 0);
    }));
    // All bits set for equal lanes
    assert!(r.iter().all(|&b| b == 0xFF));
}

#[test]
fn f64x2_sqrt() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(9.0));
        c.emit_op(Op::F64X2_SQRT, 0);
    }));
    assert_eq!(f64_lanes(&r), [3.0, 3.0]);
}

#[test]
fn f64x2_abs_neg() {
    let mk = |v: f64| {
        let mut b = [0u8; 16];
        b[0..8].copy_from_slice(&v.to_le_bytes());
        b[8..16].copy_from_slice(&v.to_le_bytes());
        b
    };
    let r_abs = as_v128(run(|c| {
        emit_v128_const(c, mk(-3.0));
        c.emit_op(Op::F64X2_ABS, 0);
    }));
    assert_eq!(f64_lanes(&r_abs), [3.0, 3.0]);
    let r_neg = as_v128(run(|c| {
        emit_v128_const(c, mk(3.0));
        c.emit_op(Op::F64X2_NEG, 0);
    }));
    assert_eq!(f64_lanes(&r_neg), [-3.0, -3.0]);
}

#[test]
fn f64x2_extract_lane() {
    let mut bytes = [0u8; 16];
    bytes[8..16].copy_from_slice(&7.0f64.to_le_bytes());
    let r = run(|c| {
        emit_v128_const(c, bytes);
        c.emit_op(Op::F64X2_EXTRACT_LANE, 0);
        c.emit(1u8, 0);
    });
    assert_eq!(r.as_f64(), 7.0);
}

#[test]
fn f64x2_replace_lane() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, [0; 16]);
        push_f64(c, 9.0);
        c.emit_op(Op::F64X2_REPLACE_LANE, 0);
        c.emit(0u8, 0);
    }));
    assert_eq!(f64_lanes(&r)[0], 9.0);
}

// ── v128 extended load variants ───────────────────────────────────────────

fn mem_run(setup: impl FnOnce(&mut VM), emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut vm = VM::new();
    vm.memory.resize(64, 0);
    setup(&mut vm);
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    vm.run(vec![c]).expect("run failed")
}

fn simd_mem_err(mem_size: usize, emit: impl FnOnce(&mut Chunk)) -> String {
    let mut vm = VM::new();
    vm.memory.resize(mem_size, 0);
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    vm.run(vec![c]).unwrap_err().to_string()
}

fn simd_store_lane_memory(op: Op, lane: u8, addr: i32, vec: [u8; 16], mem_size: usize) -> Vec<u8> {
    let mut vm = VM::new();
    vm.memory.resize(mem_size, 0);
    let mut c = Chunk::new("<script>");
    emit_v128_const(&mut c, vec);
    push_i32(&mut c, addr);
    c.emit_op(op, 0);
    c.emit(lane, 0);
    push_i32(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    vm.run(vec![c]).expect("run failed");
    let mut out = vec![0; mem_size];
    vm.memory.read_bytes(0, &mut out);
    out
}

fn assert_simd_oob(mem_size: usize, emit: impl FnOnce(&mut Chunk)) {
    let err = simd_mem_err(mem_size, emit);
    assert!(
        err.contains("out of bounds") || err.contains("trap"),
        "expected SIMD memory trap, got {err}"
    );
}

#[test]
fn v128_load_oob_traps() {
    assert_simd_oob(15, |c| {
        push_i32(c, 0);
        c.emit_op(Op::V128_LOAD, 0);
    });
}

#[test]
fn v128_store_oob_traps() {
    assert_simd_oob(15, |c| {
        push_i32(c, 0);
        emit_v128_const(c, [1; 16]);
        c.emit_op(Op::V128_STORE, 0);
    });
}

#[test]
fn v128_load64_lane_oob_traps() {
    assert_simd_oob(7, |c| {
        push_i32(c, 0);
        emit_v128_const(c, [0; 16]);
        c.emit_op(Op::V128_LOAD64_LANE, 0);
        c.emit(0u8, 0);
    });
}

#[test]
fn simd_load_extend_and_splat_variants_oob_trap() {
    let cases: &[(Op, usize)] = &[
        (Op::V128_LOAD8X8_S, 7),
        (Op::V128_LOAD8X8_U, 7),
        (Op::V128_LOAD16X4_S, 7),
        (Op::V128_LOAD16X4_U, 7),
        (Op::V128_LOAD32X2_S, 7),
        (Op::V128_LOAD32X2_U, 7),
        (Op::V128_LOAD8_SPLAT, 0),
        (Op::V128_LOAD16_SPLAT, 1),
        (Op::V128_LOAD32_SPLAT, 3),
        (Op::V128_LOAD64_SPLAT, 7),
        (Op::V128_LOAD32_ZERO, 3),
        (Op::V128_LOAD64_ZERO, 7),
    ];

    for (op, mem_size) in cases {
        assert_simd_oob(*mem_size, |c| {
            push_i32(c, 0);
            c.emit_op(*op, 0);
        });
    }
}

#[test]
fn simd_load_lane_variants_oob_trap() {
    let cases: &[(Op, usize)] = &[
        (Op::V128_LOAD8_LANE, 0),
        (Op::V128_LOAD16_LANE, 1),
        (Op::V128_LOAD32_LANE, 3),
        (Op::V128_LOAD64_LANE, 7),
    ];

    for (op, mem_size) in cases {
        assert_simd_oob(*mem_size, |c| {
            push_i32(c, 0);
            emit_v128_const(c, [0; 16]);
            c.emit_op(*op, 0);
            c.emit(0u8, 0);
        });
    }
}

#[test]
fn simd_store_lane_variants_oob_trap() {
    let cases: &[(Op, usize)] = &[
        (Op::V128_STORE8_LANE, 0),
        (Op::V128_STORE16_LANE, 1),
        (Op::V128_STORE32_LANE, 3),
        (Op::V128_STORE64_LANE, 7),
    ];

    for (op, mem_size) in cases {
        assert_simd_oob(*mem_size, |c| {
            emit_v128_const(c, [1; 16]);
            push_i32(c, 0);
            c.emit_op(*op, 0);
            c.emit(0u8, 0);
            push_i32(c, 0);
        });
    }
}

#[test]
fn v128_load8x8_s_sign_extends() {
    // Write 0xFF (-1 as i8) to addr 0; expect sign-extended to -1 as i16 in all 8 lanes
    let r = as_v128(mem_run(
        |vm| {
            for i in 0..8 {
                let _ = vm.memory.store_u8(i, 0xFF);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD8X8_S, 0);
        },
    ));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), -1, "lane {i}");
    }
}

#[test]
fn v128_load8x8_u_zero_extends() {
    let r = as_v128(mem_run(
        |vm| {
            for i in 0..8 {
                let _ = vm.memory.store_u8(i, 200);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD8X8_U, 0);
        },
    ));
    for i in 0..8 {
        assert_eq!(
            u16::from_le_bytes([r[i * 2], r[i * 2 + 1]]),
            200,
            "lane {i}"
        );
    }
}

#[test]
fn v128_load16x4_s_sign_extends() {
    let r = as_v128(mem_run(
        |vm| {
            for i in 0..4 {
                let _ = vm.memory.store_u8(i * 2, 0xFF);
                let _ = vm.memory.store_u8(i * 2 + 1, 0xFF);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD16X4_S, 0);
        },
    ));
    for i in 0..4 {
        assert_eq!(
            i32::from_le_bytes([r[i * 4], r[i * 4 + 1], r[i * 4 + 2], r[i * 4 + 3]]),
            -1,
            "lane {i}"
        );
    }
}

#[test]
fn v128_load16x4_u_zero_extends() {
    let r = as_v128(mem_run(
        |vm| {
            for i in 0..4 {
                let _ = vm.memory.store_u8(i * 2, 0x00);
                let _ = vm.memory.store_u8(i * 2 + 1, 0x01);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD16X4_U, 0);
        },
    ));
    for i in 0..4 {
        assert_eq!(
            i32::from_le_bytes([r[i * 4], r[i * 4 + 1], r[i * 4 + 2], r[i * 4 + 3]]),
            256,
            "lane {i}"
        );
    }
}

#[test]
fn v128_load32x2_s_sign_extends() {
    let r = as_v128(mem_run(
        |vm| {
            for j in 0..8 {
                let _ = vm.memory.store_u8(j, 0xFF);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD32X2_S, 0);
        },
    ));
    for i in 0..2 {
        assert_eq!(
            i64::from_le_bytes(r[i * 8..i * 8 + 8].try_into().unwrap()),
            -1i64,
            "lane {i}"
        );
    }
}

#[test]
fn v128_load32x2_u_zero_extends() {
    let r = as_v128(mem_run(
        |vm| {
            for j in 0..8 {
                let _ = vm.memory.store_u8(j, 0xFF);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD32X2_U, 0);
        },
    ));
    for i in 0..2 {
        assert_eq!(
            i64::from_le_bytes(r[i * 8..i * 8 + 8].try_into().unwrap()),
            0xFFFF_FFFFi64,
            "lane {i}"
        );
    }
}

#[test]
fn v128_load8_splat() {
    let r = as_v128(mem_run(
        |vm| {
            let _ = vm.memory.store_u8(0, 42);
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD8_SPLAT, 0);
        },
    ));
    assert!(r.iter().all(|&b| b == 42));
}

#[test]
fn v128_load16_splat() {
    let r = as_v128(mem_run(
        |vm| {
            let _ = vm.memory.store_u8(0, 0x34);
            let _ = vm.memory.store_u8(1, 0x12);
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD16_SPLAT, 0);
        },
    ));
    for i in 0..8 {
        assert_eq!(
            u16::from_le_bytes([r[i * 2], r[i * 2 + 1]]),
            0x1234,
            "lane {i}"
        );
    }
}

#[test]
fn v128_load32_splat() {
    let r = as_v128(mem_run(
        |vm| {
            let bytes = 7i32.to_le_bytes();
            for j in 0..4 {
                let _ = vm.memory.store_u8(j, bytes[j]);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD32_SPLAT, 0);
        },
    ));
    assert_eq!(i32_lanes(&r), [7, 7, 7, 7]);
}

#[test]
fn v128_load64_splat() {
    let r = as_v128(mem_run(
        |vm| {
            let bytes = 99i64.to_le_bytes();
            for j in 0..8 {
                let _ = vm.memory.store_u8(j, bytes[j]);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD64_SPLAT, 0);
        },
    ));
    let lo = i64::from_le_bytes(r[0..8].try_into().unwrap());
    let hi = i64::from_le_bytes(r[8..16].try_into().unwrap());
    assert_eq!(lo, 99);
    assert_eq!(hi, 99);
}

#[test]
fn v128_load32_zero() {
    let r = as_v128(mem_run(
        |vm| {
            let bytes = 55i32.to_le_bytes();
            for j in 0..4 {
                let _ = vm.memory.store_u8(j, bytes[j]);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD32_ZERO, 0);
        },
    ));
    assert_eq!(i32::from_le_bytes([r[0], r[1], r[2], r[3]]), 55);
    assert_eq!(&r[4..16], &[0u8; 12]);
}

#[test]
fn v128_load64_zero() {
    let r = as_v128(mem_run(
        |vm| {
            let bytes = 123i64.to_le_bytes();
            for j in 0..8 {
                let _ = vm.memory.store_u8(j, bytes[j]);
            }
        },
        |c| {
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD64_ZERO, 0);
        },
    ));
    assert_eq!(i64::from_le_bytes(r[0..8].try_into().unwrap()), 123);
    assert_eq!(&r[8..16], &[0u8; 8]);
}

// ── load_lane / store_lane ────────────────────────────────────────────────

#[test]
fn v128_load8_lane_replaces_one_byte() {
    let r = as_v128(mem_run(
        |vm| {
            let _ = vm.memory.store_u8(0, 77);
        },
        |c| {
            push_i32(c, 0); // addr
            emit_v128_const(c, [0; 16]); // vec
            c.emit_op(Op::V128_LOAD8_LANE, 0);
            c.emit(5u8, 0); // lane 5
        },
    ));
    assert_eq!(r[5], 77);
    assert!(
        r.iter()
            .enumerate()
            .filter(|&(i, _)| i != 5)
            .all(|(_, &b)| b == 0)
    );
}

#[test]
fn v128_load16_lane_replaces_two_bytes() {
    let r = as_v128(mem_run(
        |vm| {
            let _ = vm.memory.store_u8(2, 0x34);
            let _ = vm.memory.store_u8(3, 0x12);
        },
        |c| {
            push_i32(c, 2);
            emit_v128_const(c, [0xAA; 16]);
            c.emit_op(Op::V128_LOAD16_LANE, 0);
            c.emit(4u8, 0);
        },
    ));

    assert_eq!(&r[8..10], &[0x34, 0x12]);
    assert_eq!(&r[0..8], &[0xAA; 8]);
    assert_eq!(&r[10..16], &[0xAA; 6]);
}

#[test]
fn v128_load32_lane_replaces_four_bytes() {
    let r = as_v128(mem_run(
        |vm| {
            for (i, byte) in 0x1234_5678u32.to_le_bytes().iter().enumerate() {
                let _ = vm.memory.store_u8(4 + i, *byte);
            }
        },
        |c| {
            push_i32(c, 4);
            emit_v128_const(c, [0xAA; 16]);
            c.emit_op(Op::V128_LOAD32_LANE, 0);
            c.emit(2u8, 0);
        },
    ));

    assert_eq!(&r[8..12], &0x1234_5678u32.to_le_bytes());
    assert_eq!(&r[0..8], &[0xAA; 8]);
    assert_eq!(&r[12..16], &[0xAA; 4]);
}

#[test]
fn v128_store8_lane_writes_one_byte() {
    let mut vec = [0u8; 16];
    vec[3] = 99;
    let memory = simd_store_lane_memory(Op::V128_STORE8_LANE, 3, 10, vec, 16);
    assert_eq!(memory[10], 99);
    assert!(memory.iter().enumerate().all(|(i, &b)| i == 10 || b == 0));
}

#[test]
fn v128_store16_lane_writes_two_bytes() {
    let mut vec = [0u8; 16];
    vec[4..6].copy_from_slice(&0x1234u16.to_le_bytes());
    let memory = simd_store_lane_memory(Op::V128_STORE16_LANE, 2, 10, vec, 16);
    assert_eq!(&memory[10..12], &0x1234u16.to_le_bytes());
    assert!(
        memory
            .iter()
            .enumerate()
            .all(|(i, &b)| (10..12).contains(&i) || b == 0)
    );
}

#[test]
fn v128_store32_lane_writes_four_bytes() {
    let mut vec = [0u8; 16];
    vec[4..8].copy_from_slice(&0x1234_5678u32.to_le_bytes());
    let memory = simd_store_lane_memory(Op::V128_STORE32_LANE, 1, 8, vec, 16);
    assert_eq!(&memory[8..12], &0x1234_5678u32.to_le_bytes());
    assert!(
        memory
            .iter()
            .enumerate()
            .all(|(i, &b)| (8..12).contains(&i) || b == 0)
    );
}

#[test]
fn v128_store64_lane_writes_eight_bytes() {
    let mut vec = [0u8; 16];
    vec[8..16].copy_from_slice(&0x0102_0304_0506_0708u64.to_le_bytes());
    let memory = simd_store_lane_memory(Op::V128_STORE64_LANE, 1, 4, vec, 16);
    assert_eq!(&memory[4..12], &0x0102_0304_0506_0708u64.to_le_bytes());
    assert!(
        memory
            .iter()
            .enumerate()
            .all(|(i, &b)| (4..12).contains(&i) || b == 0)
    );
}

#[test]
fn v128_load64_lane_replaces_upper_lane() {
    let r = as_v128(mem_run(
        |vm| {
            let bytes = 0xDEADBEEFi64.to_le_bytes();
            for j in 0..8 {
                let _ = vm.memory.store_u8(j, bytes[j]);
            }
        },
        |c| {
            push_i32(c, 0); // addr
            emit_v128_const(c, [0; 16]); // vec
            c.emit_op(Op::V128_LOAD64_LANE, 0);
            c.emit(1u8, 0); // lane 1 (upper 8 bytes)
        },
    ));
    assert_eq!(
        i64::from_le_bytes(r[8..16].try_into().unwrap()),
        0xDEADBEEFi64
    );
    assert_eq!(&r[0..8], &[0u8; 8]);
}

// ── i8x16 unary & comparison extensions ──────────────────────────────────

#[test]
fn i8x16_abs() {
    let mut a = [0u8; 16];
    a[0] = 0xFF; // -1 as i8
    a[1] = 5;
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I8X16_ABS, 0);
    }));
    assert_eq!(r[0], 1); // abs(-1) = 1
    assert_eq!(r[1], 5);
}

#[test]
fn i8x16_neg() {
    let mut a = [0u8; 16];
    a[0] = 1;
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I8X16_NEG, 0);
    }));
    assert_eq!(r[0], 0xFF); // -1 as u8
}

#[test]
fn i8x16_popcnt() {
    let mut a = [0u8; 16];
    a[0] = 0b10110011; // 5 bits set
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I8X16_POPCNT, 0);
    }));
    assert_eq!(r[0], 5);
    assert_eq!(r[1], 0);
}

#[test]
fn i8x16_all_true_all_nonzero() {
    let r = run(|c| {
        emit_v128_const(c, [1; 16]);
        c.emit_op(Op::I8X16_ALL_TRUE, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn i8x16_all_true_with_zero() {
    let mut a = [1u8; 16];
    a[8] = 0;
    let r = run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I8X16_ALL_TRUE, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn i8x16_bitmask() {
    // All 0xFF (-1) → all sign bits set → bitmask = 0xFFFF
    let r = run(|c| {
        emit_v128_const(c, [0xFF; 16]);
        c.emit_op(Op::I8X16_BITMASK, 0);
    });
    assert_eq!(r.as_i32() as u32, 0xFFFF);
}

#[test]
fn i8x16_ne() {
    let r = as_v128(run(|c| {
        let mut a = [1u8; 16];
        a[0] = 2;
        emit_v128_const(c, a);
        emit_v128_const(c, [1u8; 16]);
        c.emit_op(Op::I8X16_NE, 0);
    }));
    assert_eq!(r[0], 0xFF); // lane 0 differs
    assert_eq!(r[1], 0x00); // lane 1 equal
}

#[test]
fn i8x16_lt_s() {
    let mut a = [0u8; 16];
    a[0] = 0xFF; // -1
    let mut b = [0u8; 16];
    b[0] = 0; // 0; -1 < 0
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_LT_S, 0);
    }));
    assert_eq!(r[0], 0xFF);
}

#[test]
fn i8x16_gt_s() {
    let mut a = [0u8; 16];
    a[0] = 1;
    let mut b = [0u8; 16];
    b[0] = 0xFF; // -1; 1 > -1
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_GT_S, 0);
    }));
    assert_eq!(r[0], 0xFF);
}

#[test]
fn i8x16_le_u() {
    let a = [5u8; 16];
    let b = [5u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_LE_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF)); // 5 <= 5
}

#[test]
fn i8x16_ge_s() {
    let a = [10u8; 16];
    let b = [5u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_GE_S, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF)); // 10 >= 5
}

// ── i8x16 min/max/avgr/sat ────────────────────────────────────────────────

#[test]
fn i8x16_min_s() {
    let mut a = [0u8; 16];
    a[0] = 0xFF; // -1
    let b = [0u8; 16]; // 0; min(-1,0) = -1
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_MIN_S, 0);
    }));
    assert_eq!(r[0], 0xFF);
}

#[test]
fn i8x16_max_u() {
    let a = [3u8; 16];
    let b = [200u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_MAX_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 200));
}

#[test]
fn i8x16_avgr_u() {
    let a = [10u8; 16];
    let b = [20u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_AVGR_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 15)); // (10+20+1)/2 = 15
}

#[test]
fn i8x16_add_sat_s_saturates() {
    let mut a = [0u8; 16];
    a[0] = 127; // i8::MAX
    let b = [1u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_ADD_SAT_S, 0);
    }));
    assert_eq!(r[0], 127); // saturated
}

#[test]
fn i8x16_sub_sat_u_saturates() {
    let a = [0u8; 16];
    let b = [1u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_SUB_SAT_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 0)); // saturated at 0
}

// ── i8x16 shifts ─────────────────────────────────────────────────────────

#[test]
fn i8x16_shl() {
    let a = [1u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        push_i32(c, 3);
        c.emit_op(Op::I8X16_SHL, 0);
    }));
    assert!(r.iter().all(|&b| b == 8));
}

#[test]
fn i8x16_shr_s_arithmetic() {
    let a = [0xFF; 16]; // -1 as i8
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        push_i32(c, 1);
        c.emit_op(Op::I8X16_SHR_S, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF)); // -1 >> 1 = -1 (arithmetic)
}

#[test]
fn i8x16_shr_u_logical() {
    let a = [0x80; 16]; // 128 unsigned
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        push_i32(c, 1);
        c.emit_op(Op::I8X16_SHR_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 0x40)); // 64
}

// ── i8x16 narrow ─────────────────────────────────────────────────────────

#[test]
fn i8x16_narrow_i16x8_s() {
    // Each i16 lane = 256 → clamps to 127 signed
    let mut a = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&256i16.to_le_bytes());
    }
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, a);
        c.emit_op(Op::I8X16_NARROW_I16X8_S, 0);
    }));
    assert!(r.iter().all(|&b| b == 127));
}

#[test]
fn i8x16_narrow_i16x8_u() {
    let mut a = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&300i16.to_le_bytes());
    }
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, a);
        c.emit_op(Op::I8X16_NARROW_I16X8_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 255)); // clamped to 255 unsigned
}

// ── i16x8 unary/comparison/all_true/bitmask ───────────────────────────────

#[test]
fn i16x8_abs_neg() {
    let mut a = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&(-5i16).to_le_bytes());
    }
    let r_abs = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_ABS, 0);
    }));
    for i in 0..8 {
        assert_eq!(
            i16::from_le_bytes([r_abs[i * 2], r_abs[i * 2 + 1]]),
            5,
            "abs lane {i}"
        );
    }
    let r_neg = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_NEG, 0);
    }));
    for i in 0..8 {
        assert_eq!(
            i16::from_le_bytes([r_neg[i * 2], r_neg[i * 2 + 1]]),
            5,
            "neg lane {i}"
        );
    }
}

#[test]
fn i16x8_all_true_and_bitmask() {
    let mut a = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&1i16.to_le_bytes());
    }
    assert_eq!(
        run(|c| {
            emit_v128_const(c, a);
            c.emit_op(Op::I16X8_ALL_TRUE, 0);
        })
        .as_i32(),
        1
    );
    // sign bits: all zeros (positive) → bitmask = 0
    assert_eq!(
        run(|c| {
            emit_v128_const(c, a);
            c.emit_op(Op::I16X8_BITMASK, 0);
        })
        .as_i32(),
        0
    );
    // Negative values → all sign bits set
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&(-1i16).to_le_bytes());
    }
    assert_eq!(
        run(|c| {
            emit_v128_const(c, a);
            c.emit_op(Op::I16X8_BITMASK, 0);
        })
        .as_i32() as u32,
        0xFF
    );
}

#[test]
fn i16x8_eq_ne() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let eq = as_v128(run(|c| {
        emit_v128_const(c, mk(7));
        emit_v128_const(c, mk(7));
        c.emit_op(Op::I16X8_EQ, 0);
    }));
    assert!(eq.iter().all(|&b| b == 0xFF));
    let ne = as_v128(run(|c| {
        emit_v128_const(c, mk(7));
        emit_v128_const(c, mk(8));
        c.emit_op(Op::I16X8_NE, 0);
    }));
    assert!(ne.iter().all(|&b| b == 0xFF));
}

#[test]
fn i16x8_lt_s_and_gt_u() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let lt = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        emit_v128_const(c, mk(0));
        c.emit_op(Op::I16X8_LT_S, 0);
    }));
    assert!(lt.iter().all(|&b| b == 0xFF)); // -1 < 0
    let gt = as_v128(run(|c| {
        emit_v128_const(c, mk(-1i16 as i16));
        emit_v128_const(c, mk(0));
        c.emit_op(Op::I16X8_GT_U, 0);
    }));
    assert!(gt.iter().all(|&b| b == 0xFF)); // 0xFFFF > 0 unsigned
}

#[test]
fn i16x8_shifts() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let shl = as_v128(run(|c| {
        emit_v128_const(c, mk(1));
        push_i32(c, 2);
        c.emit_op(Op::I16X8_SHL, 0);
    }));
    for i in 0..8 {
        assert_eq!(
            i16::from_le_bytes([shl[i * 2], shl[i * 2 + 1]]),
            4,
            "shl lane {i}"
        );
    }
    let shr_s = as_v128(run(|c| {
        emit_v128_const(c, mk(-8));
        push_i32(c, 1);
        c.emit_op(Op::I16X8_SHR_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(
            i16::from_le_bytes([shr_s[i * 2], shr_s[i * 2 + 1]]),
            -4,
            "shr_s lane {i}"
        );
    }
    let shr_u = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        push_i32(c, 1);
        c.emit_op(Op::I16X8_SHR_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(
            u16::from_le_bytes([shr_u[i * 2], shr_u[i * 2 + 1]]),
            0x7FFF,
            "shr_u lane {i}"
        );
    }
}

#[test]
fn i16x8_sub() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(10));
        emit_v128_const(c, mk(3));
        c.emit_op(Op::I16X8_SUB, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 7);
    }
}

#[test]
fn i16x8_sat_arithmetic() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let add = as_v128(run(|c| {
        emit_v128_const(c, mk(i16::MAX));
        emit_v128_const(c, mk(1));
        c.emit_op(Op::I16X8_ADD_SAT_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([add[i * 2], add[i * 2 + 1]]), i16::MAX);
    }
    let sub = as_v128(run(|c| {
        emit_v128_const(c, mk(0));
        emit_v128_const(c, mk(1));
        c.emit_op(Op::I16X8_SUB_SAT_U, 0);
    }));
    assert!(sub.iter().all(|&b| b == 0));
}

#[test]
fn i16x8_min_max_avgr() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let min = as_v128(run(|c| {
        emit_v128_const(c, mk(-5));
        emit_v128_const(c, mk(5));
        c.emit_op(Op::I16X8_MIN_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([min[i * 2], min[i * 2 + 1]]), -5);
    }
    let max = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(7));
        c.emit_op(Op::I16X8_MAX_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([max[i * 2], max[i * 2 + 1]]), 7);
    }
    let avgr = as_v128(run(|c| {
        emit_v128_const(c, mk(10));
        emit_v128_const(c, mk(11));
        c.emit_op(Op::I16X8_AVGR_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([avgr[i * 2], avgr[i * 2 + 1]]), 11);
    } // (10+11+1)/2=11
}

#[test]
fn i16x8_narrow_i32x4_s() {
    let mut a = [0u8; 16];
    for i in 0..4 {
        a[i * 4..i * 4 + 4].copy_from_slice(&40000i32.to_le_bytes());
    } // > i16::MAX
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_NARROW_I32X4_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), i16::MAX);
    }
}

// ── i16x8 extend & extmul ────────────────────────────────────────────────

#[test]
fn i16x8_extend_low_high_i8x16_s() {
    let mut a = [0u8; 16];
    a[0] = 0xFF; // -1 at lane 0
    let lo = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_EXTEND_LOW_I8X16_S, 0);
    }));
    assert_eq!(i16::from_le_bytes([lo[0], lo[1]]), -1);
    let hi = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_EXTEND_HIGH_I8X16_S, 0);
    }));
    assert!(hi.iter().all(|&b| b == 0)); // high 8 bytes of a are all zero
}

#[test]
fn i16x8_extmul_low_i8x16_s() {
    let mut a = [0u8; 16];
    let mut b = [0u8; 16];
    a[0] = 3;
    b[0] = 4;
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I16X8_EXTMUL_LOW_I8X16_S, 0);
    }));
    assert_eq!(i16::from_le_bytes([r[0], r[1]]), 12);
}

// ── i32x4 abs/neg/all_true/bitmask ───────────────────────────────────────

#[test]
fn i32x4_abs_neg() {
    let a: [u8; 16] = {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&(-3i32).to_le_bytes());
        }
        b
    };
    let r_abs = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_ABS, 0);
    }));
    assert_eq!(i32_lanes(&r_abs), [3, 3, 3, 3]);
    let r_neg = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_NEG, 0);
    }));
    assert_eq!(i32_lanes(&r_neg), [3, 3, 3, 3]);
}

#[test]
fn i32x4_all_true_and_bitmask() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    assert_eq!(
        run(|c| {
            emit_v128_const(c, mk(1));
            c.emit_op(Op::I32X4_ALL_TRUE, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run(|c| {
            emit_v128_const(c, mk(-1));
            c.emit_op(Op::I32X4_BITMASK, 0);
        })
        .as_i32() as u32,
        0xF
    );
}

#[test]
fn i32x4_comparisons() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let ne = as_v128(run(|c| {
        emit_v128_const(c, mk(1));
        emit_v128_const(c, mk(2));
        c.emit_op(Op::I32X4_NE, 0);
    }));
    assert_eq!(i32_lanes(&ne), [-1, -1, -1, -1]);
    let lt = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        emit_v128_const(c, mk(0));
        c.emit_op(Op::I32X4_LT_S, 0);
    }));
    assert_eq!(i32_lanes(&lt), [-1, -1, -1, -1]);
    let gt = as_v128(run(|c| {
        emit_v128_const(c, mk(1));
        emit_v128_const(c, mk(0));
        c.emit_op(Op::I32X4_GT_U, 0);
    }));
    assert_eq!(i32_lanes(&gt), [-1, -1, -1, -1]);
}

#[test]
fn i32x4_min_max() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let mn = as_v128(run(|c| {
        emit_v128_const(c, mk(-5));
        emit_v128_const(c, mk(5));
        c.emit_op(Op::I32X4_MIN_S, 0);
    }));
    assert_eq!(i32_lanes(&mn), [-5, -5, -5, -5]);
    let mx = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(7));
        c.emit_op(Op::I32X4_MAX_U, 0);
    }));
    assert_eq!(i32_lanes(&mx), [7, 7, 7, 7]);
}

#[test]
fn i32x4_extend_low_high_i16x8_s() {
    let mut a = [0u8; 16];
    for i in 0..4 {
        a[i * 2..i * 2 + 2].copy_from_slice(&(-1i16).to_le_bytes());
    } // low 4 lanes = -1
    let lo = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_EXTEND_LOW_I16X8_S, 0);
    }));
    assert_eq!(i32_lanes(&lo), [-1, -1, -1, -1]);
    let hi = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_EXTEND_HIGH_I16X8_S, 0);
    }));
    assert_eq!(i32_lanes(&hi), [0, 0, 0, 0]);
}

#[test]
fn i32x4_extmul_and_dot() {
    let mk16 = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk16(3));
        emit_v128_const(c, mk16(4));
        c.emit_op(Op::I32X4_EXTMUL_LOW_I16X8_S, 0);
    }));
    assert_eq!(i32_lanes(&r), [12, 12, 12, 12]);
    // dot: pairs of i16 multiplied and summed → i32 lanes
    // [3,3, 3,3, 3,3, 3,3] dot [4,4, 4,4, 4,4, 4,4] = [3*4+3*4, ...] = [24,24,24,24]
    let dot = as_v128(run(|c| {
        emit_v128_const(c, mk16(3));
        emit_v128_const(c, mk16(4));
        c.emit_op(Op::I32X4_DOT_I16X8_S, 0);
    }));
    assert_eq!(i32_lanes(&dot), [24, 24, 24, 24]);
}

#[test]
fn i32x4_extadd_pairwise() {
    let mut a = [0u8; 16];
    for i in 0..16 {
        a[i] = 1;
    } // all i8 lanes = 1
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_EXTADD_PAIRWISE_I8X16_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 2, "lane {i}");
    }

    let mut b = [0u8; 16];
    for i in 0..8 {
        b[i * 2..i * 2 + 2].copy_from_slice(&1i16.to_le_bytes());
    }
    let r2 = as_v128(run(|c| {
        emit_v128_const(c, b);
        c.emit_op(Op::I32X4_EXTADD_PAIRWISE_I16X8_S, 0);
    }));
    assert_eq!(i32_lanes(&r2), [2, 2, 2, 2]);
}

// ── i64x2 ────────────────────────────────────────────────────────────────

fn i64_lanes(b: &[u8; 16]) -> [i64; 2] {
    [
        i64::from_le_bytes(b[0..8].try_into().unwrap()),
        i64::from_le_bytes(b[8..16].try_into().unwrap()),
    ]
}

fn mk_i64x2(v: i64) -> [u8; 16] {
    let mut b = [0u8; 16];
    b[0..8].copy_from_slice(&v.to_le_bytes());
    b[8..16].copy_from_slice(&v.to_le_bytes());
    b
}

#[test]
fn i64x2_splat_extract_replace() {
    let k = c_const_i64(7i64);
    let r_splat = as_v128(run(|c| {
        let k = c.add_constant(Value::I64(7));
        c.emit_op_u16(Op::CONST, k, 0);
        c.emit_op(Op::I64X2_SPLAT, 0);
    }));
    assert_eq!(i64_lanes(&r_splat), [7, 7]);

    let r_extract = run(|c| {
        emit_v128_const(c, mk_i64x2(99));
        c.emit_op(Op::I64X2_EXTRACT_LANE, 0);
        c.emit(1u8, 0);
    });
    assert_eq!(r_extract.as_i64(), 99);

    let r_replace = as_v128(run(|c| {
        emit_v128_const(c, [0; 16]);
        let k = c.add_constant(Value::I64(55));
        c.emit_op_u16(Op::CONST, k, 0);
        c.emit_op(Op::I64X2_REPLACE_LANE, 0);
        c.emit(0u8, 0);
    }));
    assert_eq!(i64_lanes(&r_replace), [55, 0]);

    let _ = k;
}

fn c_const_i64(_v: i64) {} // dummy to allow use before definition

#[test]
fn i64x2_arithmetic() {
    let r_add = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(3));
        emit_v128_const(c, mk_i64x2(4));
        c.emit_op(Op::I64X2_ADD, 0);
    }));
    assert_eq!(i64_lanes(&r_add), [7, 7]);
    let r_sub = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(10));
        emit_v128_const(c, mk_i64x2(3));
        c.emit_op(Op::I64X2_SUB, 0);
    }));
    assert_eq!(i64_lanes(&r_sub), [7, 7]);
    let r_mul = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(3));
        emit_v128_const(c, mk_i64x2(4));
        c.emit_op(Op::I64X2_MUL, 0);
    }));
    assert_eq!(i64_lanes(&r_mul), [12, 12]);
}

#[test]
fn i64x2_abs_neg() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(-5));
        c.emit_op(Op::I64X2_ABS, 0);
    }));
    assert_eq!(i64_lanes(&r), [5, 5]);
    let r2 = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(5));
        c.emit_op(Op::I64X2_NEG, 0);
    }));
    assert_eq!(i64_lanes(&r2), [-5, -5]);
}

#[test]
fn i64x2_all_true_bitmask() {
    assert_eq!(
        run(|c| {
            emit_v128_const(c, mk_i64x2(1));
            c.emit_op(Op::I64X2_ALL_TRUE, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run(|c| {
            emit_v128_const(c, mk_i64x2(-1));
            c.emit_op(Op::I64X2_BITMASK, 0);
        })
        .as_i32() as u32,
        3
    );
}

#[test]
fn i64x2_comparisons() {
    let eq = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(5));
        emit_v128_const(c, mk_i64x2(5));
        c.emit_op(Op::I64X2_EQ, 0);
    }));
    assert!(eq.iter().all(|&b| b == 0xFF));
    let ne = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(1));
        emit_v128_const(c, mk_i64x2(2));
        c.emit_op(Op::I64X2_NE, 0);
    }));
    assert!(ne.iter().all(|&b| b == 0xFF));
    let lt = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(-1));
        emit_v128_const(c, mk_i64x2(0));
        c.emit_op(Op::I64X2_LT_S, 0);
    }));
    assert!(lt.iter().all(|&b| b == 0xFF));
}

#[test]
fn i64x2_shifts() {
    let r_shl = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(1));
        push_i32(c, 3);
        c.emit_op(Op::I64X2_SHL, 0);
    }));
    assert_eq!(i64_lanes(&r_shl), [8, 8]);
    let r_shr_s = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(-8));
        push_i32(c, 1);
        c.emit_op(Op::I64X2_SHR_S, 0);
    }));
    assert_eq!(i64_lanes(&r_shr_s), [-4, -4]);
    let r_shr_u = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(8));
        push_i32(c, 1);
        c.emit_op(Op::I64X2_SHR_U, 0);
    }));
    assert_eq!(i64_lanes(&r_shr_u), [4, 4]);
}

#[test]
fn i64x2_extend_and_extmul() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let lo = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I64X2_EXTEND_LOW_I32X4_S, 0);
    }));
    assert_eq!(i64_lanes(&lo), [-1, -1]);
    let hi = as_v128(run(|c| {
        emit_v128_const(c, mk(5));
        c.emit_op(Op::I64X2_EXTEND_HIGH_I32X4_S, 0);
    }));
    assert_eq!(i64_lanes(&hi), [5, 5]);
    let em = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I64X2_EXTMUL_LOW_I32X4_S, 0);
    }));
    assert_eq!(i64_lanes(&em), [12, 12]);
}

// ── f32x4 extended ops ───────────────────────────────────────────────────

fn mk_f32x4(v: f32) -> [u8; 16] {
    let mut b = [0u8; 16];
    for i in 0..4 {
        b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
    }
    b
}

#[test]
fn f32x4_sub_div_sqrt() {
    let r_sub = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(5.0));
        emit_v128_const(c, mk_f32x4(2.0));
        c.emit_op(Op::F32X4_SUB, 0);
    }));
    assert_eq!(f32_lanes(&r_sub), [3.0, 3.0, 3.0, 3.0]);
    let r_div = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(6.0));
        emit_v128_const(c, mk_f32x4(2.0));
        c.emit_op(Op::F32X4_DIV, 0);
    }));
    assert_eq!(f32_lanes(&r_div), [3.0, 3.0, 3.0, 3.0]);
    let r_sqrt = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(9.0));
        c.emit_op(Op::F32X4_SQRT, 0);
    }));
    assert_eq!(f32_lanes(&r_sqrt), [3.0, 3.0, 3.0, 3.0]);
}

#[test]
fn f32x4_abs_neg() {
    let r_abs = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(-4.0));
        c.emit_op(Op::F32X4_ABS, 0);
    }));
    assert_eq!(f32_lanes(&r_abs), [4.0, 4.0, 4.0, 4.0]);
    let r_neg = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(4.0));
        c.emit_op(Op::F32X4_NEG, 0);
    }));
    assert_eq!(f32_lanes(&r_neg), [-4.0, -4.0, -4.0, -4.0]);
}

#[test]
fn f32x4_min_max_pmin_pmax() {
    let mk1 = || {
        let mut b = [0u8; 16];
        b[0..4].copy_from_slice(&1.0f32.to_le_bytes());
        b[4..8].copy_from_slice(&5.0f32.to_le_bytes());
        b[8..12].copy_from_slice(&1.0f32.to_le_bytes());
        b[12..16].copy_from_slice(&5.0f32.to_le_bytes());
        b
    };
    let mk2 = || {
        let mut b = [0u8; 16];
        b[0..4].copy_from_slice(&3.0f32.to_le_bytes());
        b[4..8].copy_from_slice(&2.0f32.to_le_bytes());
        b[8..12].copy_from_slice(&3.0f32.to_le_bytes());
        b[12..16].copy_from_slice(&2.0f32.to_le_bytes());
        b
    };
    let mn = as_v128(run(|c| {
        emit_v128_const(c, mk1());
        emit_v128_const(c, mk2());
        c.emit_op(Op::F32X4_MIN, 0);
    }));
    let f_mn = f32_lanes(&mn);
    assert_eq!(f_mn[0], 1.0);
    assert_eq!(f_mn[1], 2.0);
    let mx = as_v128(run(|c| {
        emit_v128_const(c, mk1());
        emit_v128_const(c, mk2());
        c.emit_op(Op::F32X4_MAX, 0);
    }));
    let f_mx = f32_lanes(&mx);
    assert_eq!(f_mx[0], 3.0);
    assert_eq!(f_mx[1], 5.0);
    let pmn = as_v128(run(|c| {
        emit_v128_const(c, mk1());
        emit_v128_const(c, mk2());
        c.emit_op(Op::F32X4_PMIN, 0);
    }));
    let f_pmn = f32_lanes(&pmn);
    assert_eq!(f_pmn[0], 1.0);
    assert_eq!(f_pmn[1], 2.0);
    let pmx = as_v128(run(|c| {
        emit_v128_const(c, mk1());
        emit_v128_const(c, mk2());
        c.emit_op(Op::F32X4_PMAX, 0);
    }));
    let f_pmx = f32_lanes(&pmx);
    assert_eq!(f_pmx[0], 3.0);
    assert_eq!(f_pmx[1], 5.0);
}

#[test]
fn f32x4_rounding() {
    let mk = |v: f32| mk_f32x4(v);
    let ceil = as_v128(run(|c| {
        emit_v128_const(c, mk(1.2));
        c.emit_op(Op::F32X4_CEIL, 0);
    }));
    assert_eq!(f32_lanes(&ceil), [2.0, 2.0, 2.0, 2.0]);
    let floor = as_v128(run(|c| {
        emit_v128_const(c, mk(1.9));
        c.emit_op(Op::F32X4_FLOOR, 0);
    }));
    assert_eq!(f32_lanes(&floor), [1.0, 1.0, 1.0, 1.0]);
    let trunc = as_v128(run(|c| {
        emit_v128_const(c, mk(1.7));
        c.emit_op(Op::F32X4_TRUNC, 0);
    }));
    assert_eq!(f32_lanes(&trunc), [1.0, 1.0, 1.0, 1.0]);
    let nearest = as_v128(run(|c| {
        emit_v128_const(c, mk(2.5));
        c.emit_op(Op::F32X4_NEAREST, 0);
    }));
    assert_eq!(f32_lanes(&nearest), [2.0, 2.0, 2.0, 2.0]); // round-to-even
}

#[test]
fn f32x4_comparisons() {
    let eq = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(3.0));
        emit_v128_const(c, mk_f32x4(3.0));
        c.emit_op(Op::F32X4_EQ, 0);
    }));
    assert!(eq.iter().all(|&b| b == 0xFF));
    let ne = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(1.0));
        emit_v128_const(c, mk_f32x4(2.0));
        c.emit_op(Op::F32X4_NE, 0);
    }));
    assert!(ne.iter().all(|&b| b == 0xFF));
    let lt = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(1.0));
        emit_v128_const(c, mk_f32x4(2.0));
        c.emit_op(Op::F32X4_LT, 0);
    }));
    assert!(lt.iter().all(|&b| b == 0xFF));
    let gt = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(2.0));
        emit_v128_const(c, mk_f32x4(1.0));
        c.emit_op(Op::F32X4_GT, 0);
    }));
    assert!(gt.iter().all(|&b| b == 0xFF));
}

// ── f64x2 extended ops ───────────────────────────────────────────────────

fn mk_f64x2(v: f64) -> [u8; 16] {
    let mut b = [0u8; 16];
    b[0..8].copy_from_slice(&v.to_le_bytes());
    b[8..16].copy_from_slice(&v.to_le_bytes());
    b
}

#[test]
fn f64x2_pmin_pmax() {
    let a = mk_f64x2(1.0);
    let b = mk_f64x2(3.0);
    let pmn = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::F64X2_PMIN, 0);
    }));
    assert_eq!(f64_lanes(&pmn), [1.0, 1.0]);
    let pmx = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::F64X2_PMAX, 0);
    }));
    assert_eq!(f64_lanes(&pmx), [3.0, 3.0]);
}

#[test]
fn f64x2_rounding() {
    let ceil = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(1.2));
        c.emit_op(Op::F64X2_CEIL, 0);
    }));
    assert_eq!(f64_lanes(&ceil), [2.0, 2.0]);
    let floor = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(1.9));
        c.emit_op(Op::F64X2_FLOOR, 0);
    }));
    assert_eq!(f64_lanes(&floor), [1.0, 1.0]);
    let trunc = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(-1.9));
        c.emit_op(Op::F64X2_TRUNC, 0);
    }));
    assert_eq!(f64_lanes(&trunc), [-1.0, -1.0]);
    let nearest = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(2.5));
        c.emit_op(Op::F64X2_NEAREST, 0);
    }));
    assert_eq!(f64_lanes(&nearest), [2.0, 2.0]); // round-to-even
}

#[test]
fn f64x2_comparisons() {
    let ne = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(1.0));
        emit_v128_const(c, mk_f64x2(2.0));
        c.emit_op(Op::F64X2_NE, 0);
    }));
    assert!(ne.iter().all(|&b| b == 0xFF));
    let lt = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(1.0));
        emit_v128_const(c, mk_f64x2(2.0));
        c.emit_op(Op::F64X2_LT, 0);
    }));
    assert!(lt.iter().all(|&b| b == 0xFF));
    let gt = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(5.0));
        emit_v128_const(c, mk_f64x2(3.0));
        c.emit_op(Op::F64X2_GT, 0);
    }));
    assert!(gt.iter().all(|&b| b == 0xFF));
    let le = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(3.0));
        emit_v128_const(c, mk_f64x2(3.0));
        c.emit_op(Op::F64X2_LE, 0);
    }));
    assert!(le.iter().all(|&b| b == 0xFF));
    let ge = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(3.0));
        emit_v128_const(c, mk_f64x2(3.0));
        c.emit_op(Op::F64X2_GE, 0);
    }));
    assert!(ge.iter().all(|&b| b == 0xFF));
}

// ── promote / demote ──────────────────────────────────────────────────────

#[test]
fn f64x2_promote_low_f32x4() {
    // low 2 f32 lanes promoted to f64x2
    let mut a = mk_f32x4(0.0);
    a[0..4].copy_from_slice(&3.0f32.to_le_bytes());
    a[4..8].copy_from_slice(&4.0f32.to_le_bytes());
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::F64X2_PROMOTE_LOW_F32X4, 0);
    }));
    let lanes = f64_lanes(&r);
    assert!((lanes[0] - 3.0).abs() < 1e-10);
    assert!((lanes[1] - 4.0).abs() < 1e-10);
}

#[test]
fn f32x4_demote_f64x2_zero() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(3.14));
        c.emit_op(Op::F32X4_DEMOTE_F64X2_ZERO, 0);
    }));
    let f = f32_lanes(&r);
    assert!((f[0] as f64 - 3.14).abs() < 1e-5);
    assert!((f[1] as f64 - 3.14).abs() < 1e-5);
    assert_eq!(f[2], 0.0); // upper lanes zeroed
    assert_eq!(f[3], 0.0);
}

// ── SIMD integer ↔ float conversions ─────────────────────────────────────

#[test]
fn i32x4_trunc_sat_f32x4_s() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(3.9));
        c.emit_op(Op::I32X4_TRUNC_SAT_F32X4_S, 0);
    }));
    assert_eq!(i32_lanes(&r), [3, 3, 3, 3]);
}

#[test]
fn i32x4_trunc_sat_f32x4_u() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(3.9));
        c.emit_op(Op::I32X4_TRUNC_SAT_F32X4_U, 0);
    }));
    assert_eq!(i32_lanes(&r).map(|v| v as u32), [3, 3, 3, 3]);
}

#[test]
fn f32x4_convert_i32x4_s() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(-4));
        c.emit_op(Op::F32X4_CONVERT_I32X4_S, 0);
    }));
    assert_eq!(f32_lanes(&r), [-4.0, -4.0, -4.0, -4.0]);
}

#[test]
fn f32x4_convert_i32x4_u() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(4));
        c.emit_op(Op::F32X4_CONVERT_I32X4_U, 0);
    }));
    assert_eq!(f32_lanes(&r), [4.0, 4.0, 4.0, 4.0]);
}

#[test]
fn i32x4_trunc_sat_f64x2_s_zero() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(7.9));
        c.emit_op(Op::I32X4_TRUNC_SAT_F64X2_S_ZERO, 0);
    }));
    assert_eq!(i32::from_le_bytes([r[0], r[1], r[2], r[3]]), 7);
    assert_eq!(i32::from_le_bytes([r[4], r[5], r[6], r[7]]), 7);
    assert_eq!(&r[8..16], &[0u8; 8]); // upper lanes zero
}

#[test]
fn f64x2_convert_low_i32x4_s() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(5));
        c.emit_op(Op::F64X2_CONVERT_LOW_I32X4_S, 0);
    }));
    assert_eq!(f64_lanes(&r), [5.0, 5.0]);
}

// ── Coverage gaps: symmetric variants not yet tested ──────────────────────

// f32x4 GE / LE
#[test]
fn f32x4_ge_le() {
    let ge = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(3.0));
        emit_v128_const(c, mk_f32x4(3.0));
        c.emit_op(Op::F32X4_GE, 0);
    }));
    assert!(ge.iter().all(|&b| b == 0xFF));
    let le = as_v128(run(|c| {
        emit_v128_const(c, mk_f32x4(2.0));
        emit_v128_const(c, mk_f32x4(3.0));
        c.emit_op(Op::F32X4_LE, 0);
    }));
    assert!(le.iter().all(|&b| b == 0xFF));
}

// f64x2 convert_low_i32x4_u (unsigned)
#[test]
fn f64x2_convert_low_i32x4_u() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::F64X2_CONVERT_LOW_I32X4_U, 0);
    }));
    assert!((f64_lanes(&r)[0] - 4_294_967_295.0).abs() < 1.0);
}

// i8x16 missing symmetrics
#[test]
fn i8x16_add_sat_u() {
    let a = [255u8; 16];
    let b = [1u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_ADD_SAT_U, 0);
    }));
    assert!(r.iter().all(|&b| b == 255)); // saturated
}

#[test]
fn i8x16_sub_sat_s() {
    let mut a = [0u8; 16];
    a[0] = 0x80; // i8::MIN
    let b = [1u8; 16];
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_SUB_SAT_S, 0);
    }));
    assert_eq!(r[0], 0x80); // saturated at i8::MIN
}

#[test]
fn i8x16_gt_u_le_s_lt_u_ge_u() {
    let lo = [1u8; 16];
    let hi = [200u8; 16]; // 200 > 1 unsigned; as signed, 200=-56 < 1
    let gt_u = as_v128(run(|c| {
        emit_v128_const(c, hi);
        emit_v128_const(c, lo);
        c.emit_op(Op::I8X16_GT_U, 0);
    }));
    assert!(gt_u.iter().all(|&b| b == 0xFF));
    let le_s = as_v128(run(|c| {
        emit_v128_const(c, hi);
        emit_v128_const(c, lo);
        c.emit_op(Op::I8X16_LE_S, 0);
    }));
    assert!(le_s.iter().all(|&b| b == 0xFF)); // -56 <= 1 signed
    let lt_u = as_v128(run(|c| {
        emit_v128_const(c, lo);
        emit_v128_const(c, hi);
        c.emit_op(Op::I8X16_LT_U, 0);
    }));
    assert!(lt_u.iter().all(|&b| b == 0xFF)); // 1 < 200 unsigned
    let ge_u = as_v128(run(|c| {
        emit_v128_const(c, hi);
        emit_v128_const(c, lo);
        c.emit_op(Op::I8X16_GE_U, 0);
    }));
    assert!(ge_u.iter().all(|&b| b == 0xFF)); // 200 >= 1 unsigned
}

#[test]
fn i8x16_max_s_min_u() {
    let a = [0u8; 16]; // 0
    let b = [0xFF; 16]; // -1 signed, 255 unsigned
    let max_s = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_MAX_S, 0);
    }));
    assert!(max_s.iter().all(|&x| x == 0)); // max(0,-1) = 0
    let min_u = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, b);
        c.emit_op(Op::I8X16_MIN_U, 0);
    }));
    assert!(min_u.iter().all(|&x| x == 0)); // min(0,255) = 0 unsigned
}

// i16x8 missing symmetrics
#[test]
fn i16x8_extract_lane_u() {
    let mut bytes = [0u8; 16];
    bytes[2..4].copy_from_slice(&0xFFFFu16.to_le_bytes());
    let r = run(|c| {
        emit_v128_const(c, bytes);
        c.emit_op(Op::I16X8_EXTRACT_LANE_U, 0);
        c.emit(1u8, 0);
    });
    assert_eq!(r.as_i32(), 65535); // unsigned, not sign-extended
}

#[test]
fn i16x8_add_sat_u() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(-1i16));
        emit_v128_const(c, mk(1));
        c.emit_op(Op::I16X8_ADD_SAT_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 0xFFFF);
    } // saturated
}

#[test]
fn i16x8_sub_sat_s() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(i16::MIN));
        emit_v128_const(c, mk(1));
        c.emit_op(Op::I16X8_SUB_SAT_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), i16::MIN);
    }
}

#[test]
fn i16x8_ge_s_ge_u_gt_s_le_s_le_u_lt_u() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let pos = mk(5);
    let neg = mk(-1);
    let big = mk(-1i16); // -1 unsigned = 0xFFFF (large)
    let ge_s = as_v128(run(|c| {
        emit_v128_const(c, pos);
        emit_v128_const(c, neg);
        c.emit_op(Op::I16X8_GE_S, 0);
    }));
    assert!(ge_s.iter().all(|&b| b == 0xFF)); // 5 >= -1
    let ge_u = as_v128(run(|c| {
        emit_v128_const(c, big);
        emit_v128_const(c, pos);
        c.emit_op(Op::I16X8_GE_U, 0);
    }));
    assert!(ge_u.iter().all(|&b| b == 0xFF)); // 0xFFFF >= 5
    let gt_s = as_v128(run(|c| {
        emit_v128_const(c, pos);
        emit_v128_const(c, neg);
        c.emit_op(Op::I16X8_GT_S, 0);
    }));
    assert!(gt_s.iter().all(|&b| b == 0xFF)); // 5 > -1
    let le_s = as_v128(run(|c| {
        emit_v128_const(c, neg);
        emit_v128_const(c, pos);
        c.emit_op(Op::I16X8_LE_S, 0);
    }));
    assert!(le_s.iter().all(|&b| b == 0xFF)); // -1 <= 5
    let le_u = as_v128(run(|c| {
        emit_v128_const(c, pos);
        emit_v128_const(c, big);
        c.emit_op(Op::I16X8_LE_U, 0);
    }));
    assert!(le_u.iter().all(|&b| b == 0xFF)); // 5 <= 0xFFFF
    let lt_u = as_v128(run(|c| {
        emit_v128_const(c, pos);
        emit_v128_const(c, big);
        c.emit_op(Op::I16X8_LT_U, 0);
    }));
    assert!(lt_u.iter().all(|&b| b == 0xFF)); // 5 < 0xFFFF
}

#[test]
fn i16x8_max_s_min_u() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let max_s = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I16X8_MAX_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([max_s[i * 2], max_s[i * 2 + 1]]), 3);
    }
    let min_u = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I16X8_MIN_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([min_u[i * 2], min_u[i * 2 + 1]]), 3);
    } // 3 < 65535
}

#[test]
fn i16x8_narrow_i32x4_u() {
    let mut a = [0u8; 16];
    for i in 0..4 {
        a[i * 4..i * 4 + 4].copy_from_slice(&70000i32.to_le_bytes());
    }
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_NARROW_I32X4_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 0xFFFF);
    } // saturated
}

#[test]
fn i16x8_q15mulr_sat_s() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    // q15mulr_sat_s: round((a*b + 0x4000) >> 15)
    // 4 * 8192 = 32768; (32768 + 16384) >> 15 = 49152 >> 15 = 1
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk(4));
        emit_v128_const(c, mk(8192));
        c.emit_op(Op::I16X8_Q15MULR_SAT_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([r[i * 2], r[i * 2 + 1]]), 1);
    }
    // saturation: i16::MIN * i16::MIN → 32769 → clamped to i16::MAX
    let sat = as_v128(run(|c| {
        emit_v128_const(c, mk(i16::MIN));
        emit_v128_const(c, mk(i16::MIN));
        c.emit_op(Op::I16X8_Q15MULR_SAT_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([sat[i * 2], sat[i * 2 + 1]]), i16::MAX);
    }
}

#[test]
fn i16x8_extadd_pairwise_u_and_extend_high_u() {
    let mut a = [0u8; 16];
    for b in a.iter_mut() {
        *b = 100;
    } // all bytes = 100
    let ep = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_EXTADD_PAIRWISE_I8X16_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([ep[i * 2], ep[i * 2 + 1]]), 200);
    }
    // extend_high_i8x16_u: upper 8 bytes → u16 lanes
    let eh = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_EXTEND_HIGH_I8X16_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([eh[i * 2], eh[i * 2 + 1]]), 100);
    }
    // extend_low_i8x16_u
    let el = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I16X8_EXTEND_LOW_I8X16_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([el[i * 2], el[i * 2 + 1]]), 100);
    }
}

#[test]
fn i16x8_extmul_variants() {
    let mk8 = |v: u8| {
        let mut b = [0u8; 16];
        for i in 0..16 {
            b[i] = v;
        }
        b
    };
    let hi_s = as_v128(run(|c| {
        emit_v128_const(c, mk8(3));
        emit_v128_const(c, mk8(4));
        c.emit_op(Op::I16X8_EXTMUL_HIGH_I8X16_S, 0);
    }));
    for i in 0..8 {
        assert_eq!(i16::from_le_bytes([hi_s[i * 2], hi_s[i * 2 + 1]]), 12);
    }
    let lo_u = as_v128(run(|c| {
        emit_v128_const(c, mk8(200));
        emit_v128_const(c, mk8(2));
        c.emit_op(Op::I16X8_EXTMUL_LOW_I8X16_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([lo_u[i * 2], lo_u[i * 2 + 1]]), 400);
    }
    let hi_u = as_v128(run(|c| {
        emit_v128_const(c, mk8(200));
        emit_v128_const(c, mk8(2));
        c.emit_op(Op::I16X8_EXTMUL_HIGH_I8X16_U, 0);
    }));
    for i in 0..8 {
        assert_eq!(u16::from_le_bytes([hi_u[i * 2], hi_u[i * 2 + 1]]), 400);
    }
}

// i32x4 missing symmetrics
#[test]
fn i32x4_ge_s_ge_u_gt_s_le_s_le_u_lt_u() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let ge_s = as_v128(run(|c| {
        emit_v128_const(c, mk(5));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_GE_S, 0);
    }));
    assert_eq!(i32_lanes(&ge_s), [-1, -1, -1, -1]);
    let ge_u = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        emit_v128_const(c, mk(5));
        c.emit_op(Op::I32X4_GE_U, 0);
    }));
    assert_eq!(i32_lanes(&ge_u), [-1, -1, -1, -1]); // 0xFFFFFFFF >= 5
    let gt_s = as_v128(run(|c| {
        emit_v128_const(c, mk(5));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_GT_S, 0);
    }));
    assert_eq!(i32_lanes(&gt_s), [-1, -1, -1, -1]);
    let le_s = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        emit_v128_const(c, mk(5));
        c.emit_op(Op::I32X4_LE_S, 0);
    }));
    assert_eq!(i32_lanes(&le_s), [-1, -1, -1, -1]);
    let le_u = as_v128(run(|c| {
        emit_v128_const(c, mk(5));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_LE_U, 0);
    }));
    assert_eq!(i32_lanes(&le_u), [-1, -1, -1, -1]); // 5 <= 0xFFFFFFFF
    let lt_u = as_v128(run(|c| {
        emit_v128_const(c, mk(5));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_LT_U, 0);
    }));
    assert_eq!(i32_lanes(&lt_u), [-1, -1, -1, -1]); // 5 < 0xFFFFFFFF
}

#[test]
fn i32x4_max_s_min_u() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let max_s = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_MAX_S, 0);
    }));
    assert_eq!(i32_lanes(&max_s), [3, 3, 3, 3]);
    let min_u = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_MIN_U, 0);
    }));
    assert_eq!(i32_lanes(&min_u), [3, 3, 3, 3]); // 3 < 0xFFFFFFFF unsigned
}

#[test]
fn i32x4_extadd_pairwise_u() {
    let mut a = [0u8; 16];
    for i in 0..8 {
        a[i * 2..i * 2 + 2].copy_from_slice(&100u16.to_le_bytes());
    }
    let r = as_v128(run(|c| {
        emit_v128_const(c, a);
        c.emit_op(Op::I32X4_EXTADD_PAIRWISE_I16X8_U, 0);
    }));
    assert_eq!(i32_lanes(&r), [200, 200, 200, 200]);
}

#[test]
fn i32x4_extend_high_u_and_low_u() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let hi = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_EXTEND_HIGH_I16X8_U, 0);
    }));
    assert_eq!(i32_lanes(&hi), [65535, 65535, 65535, 65535]); // zero-extended
    let lo = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I32X4_EXTEND_LOW_I16X8_U, 0);
    }));
    assert_eq!(i32_lanes(&lo), [65535, 65535, 65535, 65535]);
}

#[test]
fn i32x4_extmul_high_and_u_variants() {
    let mk = |v: i16| {
        let mut b = [0u8; 16];
        for i in 0..8 {
            b[i * 2..i * 2 + 2].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let hi_s = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I32X4_EXTMUL_HIGH_I16X8_S, 0);
    }));
    assert_eq!(i32_lanes(&hi_s), [12, 12, 12, 12]);
    let lo_u = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I32X4_EXTMUL_LOW_I16X8_U, 0);
    }));
    assert_eq!(i32_lanes(&lo_u), [12, 12, 12, 12]);
    let hi_u = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I32X4_EXTMUL_HIGH_I16X8_U, 0);
    }));
    assert_eq!(i32_lanes(&hi_u), [12, 12, 12, 12]);
}

#[test]
fn i32x4_trunc_sat_f64x2_u_zero() {
    let r = as_v128(run(|c| {
        emit_v128_const(c, mk_f64x2(5.9));
        c.emit_op(Op::I32X4_TRUNC_SAT_F64X2_U_ZERO, 0);
    }));
    assert_eq!(i32::from_le_bytes([r[0], r[1], r[2], r[3]]) as u32, 5);
    assert_eq!(i32::from_le_bytes([r[4], r[5], r[6], r[7]]) as u32, 5);
    assert_eq!(&r[8..16], &[0u8; 8]);
}

// i64x2 missing symmetrics
#[test]
fn i64x2_ge_s_gt_s_le_s() {
    let ge = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(5));
        emit_v128_const(c, mk_i64x2(5));
        c.emit_op(Op::I64X2_GE_S, 0);
    }));
    assert!(ge.iter().all(|&b| b == 0xFF));
    let gt = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(6));
        emit_v128_const(c, mk_i64x2(5));
        c.emit_op(Op::I64X2_GT_S, 0);
    }));
    assert!(gt.iter().all(|&b| b == 0xFF));
    let le = as_v128(run(|c| {
        emit_v128_const(c, mk_i64x2(-1));
        emit_v128_const(c, mk_i64x2(0));
        c.emit_op(Op::I64X2_LE_S, 0);
    }));
    assert!(le.iter().all(|&b| b == 0xFF));
}

#[test]
fn i64x2_extend_high_low_u_and_extmul_variants() {
    let mk = |v: i32| {
        let mut b = [0u8; 16];
        for i in 0..4 {
            b[i * 4..i * 4 + 4].copy_from_slice(&v.to_le_bytes());
        }
        b
    };
    let lo_u = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I64X2_EXTEND_LOW_I32X4_U, 0);
    }));
    assert_eq!(i64_lanes(&lo_u), [4294967295i64, 4294967295i64]);
    let hi_u = as_v128(run(|c| {
        emit_v128_const(c, mk(-1));
        c.emit_op(Op::I64X2_EXTEND_HIGH_I32X4_U, 0);
    }));
    assert_eq!(i64_lanes(&hi_u), [4294967295i64, 4294967295i64]);
    let hi_s = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I64X2_EXTMUL_HIGH_I32X4_S, 0);
    }));
    assert_eq!(i64_lanes(&hi_s), [12, 12]);
    let lo_u2 = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I64X2_EXTMUL_LOW_I32X4_U, 0);
    }));
    assert_eq!(i64_lanes(&lo_u2), [12, 12]);
    let hi_u2 = as_v128(run(|c| {
        emit_v128_const(c, mk(3));
        emit_v128_const(c, mk(4));
        c.emit_op(Op::I64X2_EXTMUL_HIGH_I32X4_U, 0);
    }));
    assert_eq!(i64_lanes(&hi_u2), [12, 12]);
}

// v128 load/store lane variants not yet tested
#[test]
fn v128_load16_lane_and_store16_lane() {
    let r = as_v128(mem_run(
        |vm| {
            let _ = vm.memory.store_u8(0, 0x34);
            let _ = vm.memory.store_u8(1, 0x12);
        },
        |c| {
            push_i32(c, 0); // addr
            emit_v128_const(c, [0; 16]); // vec
            c.emit_op(Op::V128_LOAD16_LANE, 0);
            c.emit(2u8, 0); // lane 2
        },
    ));
    assert_eq!(u16::from_le_bytes([r[4], r[5]]), 0x1234);
}

#[test]
fn v128_load32_lane_and_store32_lane() {
    let r = as_v128(mem_run(
        |vm| {
            let bytes = 0xDEAD_BEEFu32.to_le_bytes();
            for j in 0..4 {
                let _ = vm.memory.store_u8(j, bytes[j]);
            }
        },
        |c| {
            push_i32(c, 0);
            emit_v128_const(c, [0; 16]);
            c.emit_op(Op::V128_LOAD32_LANE, 0);
            c.emit(1u8, 0); // lane 1
        },
    ));
    assert_eq!(u32::from_le_bytes([r[4], r[5], r[6], r[7]]), 0xDEAD_BEEF);
}

#[test]
fn v128_store16_lane() {
    // Store lane 0 of [0x1234 as i16x8] to addr 4, then read back
    let r = as_v128(mem_run(
        |_| {},
        |c| {
            // Build v128 with 0x1234 in lane 0
            push_i32(c, 0x1234);
            c.emit_op(Op::I16X8_SPLAT, 0);
            push_i32(c, 4); // addr
            c.emit_op(Op::V128_STORE16_LANE, 0);
            c.emit(0u8, 0); // store lane 0 to addr 4
            // load it back
            push_i32(c, 4);
            c.emit_op(Op::V128_LOAD16_SPLAT, 0);
        },
    ));
    for i in 0..8 {
        assert_eq!(
            u16::from_le_bytes([r[i * 2], r[i * 2 + 1]]),
            0x1234,
            "lane {i}"
        );
    }
}

#[test]
fn v128_store32_lane() {
    let r = as_v128(mem_run(
        |_| {},
        |c| {
            push_i32(c, 42);
            c.emit_op(Op::I32X4_SPLAT, 0);
            push_i32(c, 0); // addr
            c.emit_op(Op::V128_STORE32_LANE, 0);
            c.emit(0u8, 0);
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD32_SPLAT, 0);
        },
    ));
    assert_eq!(i32_lanes(&r), [42, 42, 42, 42]);
}

#[test]
fn v128_store64_lane() {
    let r = as_v128(mem_run(
        |_| {},
        |c| {
            let k = c.add_constant(Value::I64(99));
            c.emit_op_u16(Op::CONST, k, 0);
            c.emit_op(Op::I64X2_SPLAT, 0);
            push_i32(c, 0);
            c.emit_op(Op::V128_STORE64_LANE, 0);
            c.emit(0u8, 0);
            push_i32(c, 0);
            c.emit_op(Op::V128_LOAD64_SPLAT, 0);
        },
    ));
    let lo = i64::from_le_bytes(r[0..8].try_into().unwrap());
    let hi = i64::from_le_bytes(r[8..16].try_into().unwrap());
    assert_eq!(lo, 99);
    assert_eq!(hi, 99);
}
