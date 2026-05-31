//! Tests for the relaxed-simd proposal (0xDD prefix).
//! Covers all 20 dispatched relaxed SIMD operations:
//!   i8x16.relaxed_swizzle, i32x4.relaxed_trunc_f32x4_{s,u},
//!   i32x4.relaxed_trunc_f64x2_{s,u}_zero, f32x4.relaxed_{madd,nmadd},
//!   f64x2.relaxed_{madd,nmadd}, i8x16/i16x8/i32x4/i64x2.relaxed_laneselect,
//!   f32x4/f64x2.relaxed_{min,max}, i16x8.relaxed_q15mulr_s,
//!   i16x8.relaxed_dot_i8x16_i7x16_s, i32x4.relaxed_dot_i8x16_i7x16_add_s.

use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::value::Value;

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn emit_v128(c: &mut Chunk, bytes: [u8; 16]) {
    c.emit_op(Op::V128_CONST, 0);
    for &b in &bytes { c.emit(b, 0); }
}


fn as_v128(v: Value) -> [u8; 16] {
    match v { Value::V128(b) => b, _ => panic!("expected V128, got {:?}", v) }
}

fn i32_lanes(b: &[u8; 16]) -> [i32; 4] {
    core::array::from_fn(|i| i32::from_le_bytes(b[i*4..i*4+4].try_into().unwrap()))
}

fn f32_lanes(b: &[u8; 16]) -> [f32; 4] {
    core::array::from_fn(|i| f32::from_le_bytes(b[i*4..i*4+4].try_into().unwrap()))
}

fn f64_lanes(b: &[u8; 16]) -> [f64; 2] {
    [f64::from_le_bytes(b[0..8].try_into().unwrap()), f64::from_le_bytes(b[8..16].try_into().unwrap())]
}

fn mk_i32x4(lanes: [i32; 4]) -> [u8; 16] {
    let mut b = [0u8; 16];
    for (i, v) in lanes.iter().enumerate() { b[i*4..i*4+4].copy_from_slice(&v.to_le_bytes()); }
    b
}

fn mk_f32x4(lanes: [f32; 4]) -> [u8; 16] {
    let mut b = [0u8; 16];
    for (i, v) in lanes.iter().enumerate() { b[i*4..i*4+4].copy_from_slice(&v.to_le_bytes()); }
    b
}

fn mk_f64x2(a: f64, b_val: f64) -> [u8; 16] {
    let mut buf = [0u8; 16];
    buf[0..8].copy_from_slice(&a.to_le_bytes());
    buf[8..16].copy_from_slice(&b_val.to_le_bytes());
    buf
}

// ── i8x16.relaxed_swizzle ─────────────────────────────────────────────────

#[test]
fn i8x16_relaxed_swizzle_picks_lanes() {
    let src: [u8;16] = [10,20,30,40,0,0,0,0,0,0,0,0,0,0,0,0];
    let idx: [u8;16] = [2,0,1,3,0,0,0,0,0,0,0,0,0,0,0,0];
    let r = as_v128(run(|c| {
        emit_v128(c, src); emit_v128(c, idx);
        c.emit_op(Op::I8X16_RELAXED_SWIZZLE, 0);
    }));
    assert_eq!(r[0], 30); assert_eq!(r[1], 10); assert_eq!(r[2], 20); assert_eq!(r[3], 40);
}

#[test]
fn i8x16_relaxed_swizzle_oob_index_gives_zero() {
    // Index >= 16 → 0
    let src: [u8;16] = [99;16];
    let idx: [u8;16] = [16,17,255,0,0,0,0,0,0,0,0,0,0,0,0,0];
    let r = as_v128(run(|c| { emit_v128(c, src); emit_v128(c, idx); c.emit_op(Op::I8X16_RELAXED_SWIZZLE, 0); }));
    assert_eq!(r[0], 0); assert_eq!(r[1], 0); assert_eq!(r[2], 0); assert_eq!(r[3], 99);
}

// ── i32x4.relaxed_trunc_f32x4_s ──────────────────────────────────────────

#[test]
fn i32x4_relaxed_trunc_f32x4_s() {
    let input = mk_f32x4([3.7, -2.9, 0.0, 100.0]);
    let r = i32_lanes(&as_v128(run(|c| { emit_v128(c, input); c.emit_op(Op::I32X4_RELAXED_TRUNC_F32X4_S, 0); })));
    assert_eq!(r, [3, -2, 0, 100]);
}

#[test]
fn i32x4_relaxed_trunc_f32x4_s_nan_gives_zero() {
    let input = mk_f32x4([f32::NAN, 1.0, 2.0, 3.0]);
    let r = i32_lanes(&as_v128(run(|c| { emit_v128(c, input); c.emit_op(Op::I32X4_RELAXED_TRUNC_F32X4_S, 0); })));
    assert_eq!(r[0], 0);
}

// ── i32x4.relaxed_trunc_f32x4_u ──────────────────────────────────────────

#[test]
fn i32x4_relaxed_trunc_f32x4_u() {
    let input = mk_f32x4([3.9, 0.0, 255.5, 1000.0]);
    let r = i32_lanes(&as_v128(run(|c| { emit_v128(c, input); c.emit_op(Op::I32X4_RELAXED_TRUNC_F32X4_U, 0); })));
    assert_eq!(r[0] as u32, 3); assert_eq!(r[1] as u32, 0);
    assert_eq!(r[2] as u32, 255); assert_eq!(r[3] as u32, 1000);
}

#[test]
fn i32x4_relaxed_trunc_f32x4_u_negative_gives_zero() {
    let input = mk_f32x4([-1.0, 0.0, 0.0, 0.0]);
    let r = i32_lanes(&as_v128(run(|c| { emit_v128(c, input); c.emit_op(Op::I32X4_RELAXED_TRUNC_F32X4_U, 0); })));
    assert_eq!(r[0] as u32, 0);
}

// ── i32x4.relaxed_trunc_f64x2_s_zero ─────────────────────────────────────

#[test]
fn i32x4_relaxed_trunc_f64x2_s_zero() {
    let input = mk_f64x2(3.7, -2.1);
    let r = i32_lanes(&as_v128(run(|c| { emit_v128(c, input); c.emit_op(Op::I32X4_RELAXED_TRUNC_F64X2_S_ZERO, 0); })));
    assert_eq!(r[0], 3); assert_eq!(r[1], -2); assert_eq!(r[2], 0); assert_eq!(r[3], 0);
}

// ── i32x4.relaxed_trunc_f64x2_u_zero ─────────────────────────────────────

#[test]
fn i32x4_relaxed_trunc_f64x2_u_zero() {
    let input = mk_f64x2(5.9, 100.1);
    let r = i32_lanes(&as_v128(run(|c| { emit_v128(c, input); c.emit_op(Op::I32X4_RELAXED_TRUNC_F64X2_U_ZERO, 0); })));
    assert_eq!(r[0] as u32, 5); assert_eq!(r[1] as u32, 100);
    assert_eq!(r[2], 0); assert_eq!(r[3], 0);
}

// ── f32x4.relaxed_madd / f32x4.relaxed_nmadd ─────────────────────────────

#[test]
fn f32x4_relaxed_madd() {
    // madd: a * b + c
    let a = mk_f32x4([2.0, 3.0, 1.0, 0.0]);
    let b = mk_f32x4([3.0, 4.0, 5.0, 1.0]);
    let c_v = mk_f32x4([1.0, 2.0, 3.0, 7.0]);
    let r = f32_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, c_v);
        c.emit_op(Op::F32X4_RELAXED_MADD, 0);
    })));
    assert_eq!(r, [7.0, 14.0, 8.0, 7.0]); // 2*3+1, 3*4+2, 1*5+3, 0*1+7
}

#[test]
fn f32x4_relaxed_nmadd() {
    // nmadd: -(a * b) + c
    let a = mk_f32x4([2.0, 1.0, 0.0, 0.0]);
    let b = mk_f32x4([3.0, 1.0, 0.0, 0.0]);
    let c_v = mk_f32x4([1.0, 5.0, 0.0, 0.0]);
    let r = f32_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, c_v);
        c.emit_op(Op::F32X4_RELAXED_NMADD, 0);
    })));
    assert_eq!(r[0], -5.0); // -(2*3)+1
    assert_eq!(r[1], 4.0);  // -(1*1)+5
}

// ── f64x2.relaxed_madd / f64x2.relaxed_nmadd ─────────────────────────────

#[test]
fn f64x2_relaxed_madd() {
    let a = mk_f64x2(2.0, 3.0);
    let b = mk_f64x2(4.0, 5.0);
    let c_v = mk_f64x2(1.0, 2.0);
    let r = f64_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, c_v);
        c.emit_op(Op::F64X2_RELAXED_MADD, 0);
    })));
    assert_eq!(r, [9.0, 17.0]); // 2*4+1, 3*5+2
}

#[test]
fn f64x2_relaxed_nmadd() {
    let a = mk_f64x2(2.0, 1.0);
    let b = mk_f64x2(3.0, 4.0);
    let c_v = mk_f64x2(1.0, 10.0);
    let r = f64_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, c_v);
        c.emit_op(Op::F64X2_RELAXED_NMADD, 0);
    })));
    assert_eq!(r[0], -5.0);  // -(2*3)+1
    assert_eq!(r[1], 6.0);   // -(1*4)+10
}

// ── i8x16.relaxed_laneselect ──────────────────────────────────────────────

#[test]
fn i8x16_relaxed_laneselect() {
    // laneselect: for each bit in mask, pick from a (1) or b (0)
    let a: [u8;16] = [0xAA;16]; // 10101010
    let b: [u8;16] = [0x55;16]; // 01010101
    let mask: [u8;16] = [0xFF;16];
    let r = as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, mask);
        c.emit_op(Op::I8X16_RELAXED_LANESELECT, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xAA));
}

// ── i16x8.relaxed_laneselect ──────────────────────────────────────────────

#[test]
fn i16x8_relaxed_laneselect() {
    let a: [u8;16] = [0xFF;16];
    let b: [u8;16] = [0x00;16];
    let mask: [u8;16] = [0xFF;16];
    let r = as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, mask);
        c.emit_op(Op::I16X8_RELAXED_LANESELECT, 0);
    }));
    assert!(r.iter().all(|&b| b == 0xFF));
}

// ── i32x4.relaxed_laneselect ──────────────────────────────────────────────

#[test]
fn i32x4_relaxed_laneselect() {
    let a = mk_i32x4([1, 2, 3, 4]);
    let b = mk_i32x4([5, 6, 7, 8]);
    let mask: [u8;16] = [0xFF;16]; // all bits set → pick all from a
    let r = i32_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, mask);
        c.emit_op(Op::I32X4_RELAXED_LANESELECT, 0);
    })));
    assert_eq!(r, [1, 2, 3, 4]);
}

// ── i64x2.relaxed_laneselect ──────────────────────────────────────────────

#[test]
fn i64x2_relaxed_laneselect() {
    let a: [u8;16] = [0xAA;16];
    let b: [u8;16] = [0x55;16];
    let mask: [u8;16] = [0x00;16]; // all 0 → pick all from b
    let r = as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, mask);
        c.emit_op(Op::I64X2_RELAXED_LANESELECT, 0);
    }));
    assert!(r.iter().all(|&b| b == 0x55));
}

// ── f32x4.relaxed_min / f32x4.relaxed_max ────────────────────────────────

#[test]
fn f32x4_relaxed_min() {
    let a = mk_f32x4([1.0, 5.0, 3.0, 0.0]);
    let b = mk_f32x4([2.0, 4.0, 3.0, 1.0]);
    let r = f32_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b);
        c.emit_op(Op::F32X4_RELAXED_MIN, 0);
    })));
    assert_eq!(r, [1.0, 4.0, 3.0, 0.0]);
}

#[test]
fn f32x4_relaxed_max() {
    let a = mk_f32x4([1.0, 5.0, 3.0, 0.0]);
    let b = mk_f32x4([2.0, 4.0, 3.0, 1.0]);
    let r = f32_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b);
        c.emit_op(Op::F32X4_RELAXED_MAX, 0);
    })));
    assert_eq!(r, [2.0, 5.0, 3.0, 1.0]);
}

// ── f64x2.relaxed_min / f64x2.relaxed_max ────────────────────────────────

#[test]
fn f64x2_relaxed_min() {
    let a = mk_f64x2(1.0, 5.0);
    let b = mk_f64x2(3.0, 2.0);
    let r = f64_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b);
        c.emit_op(Op::F64X2_RELAXED_MIN, 0);
    })));
    assert_eq!(r, [1.0, 2.0]);
}

#[test]
fn f64x2_relaxed_max() {
    let a = mk_f64x2(1.0, 5.0);
    let b = mk_f64x2(3.0, 2.0);
    let r = f64_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b);
        c.emit_op(Op::F64X2_RELAXED_MAX, 0);
    })));
    assert_eq!(r, [3.0, 5.0]);
}

// ── i16x8.relaxed_q15mulr_s ───────────────────────────────────────────────

#[test]
fn i16x8_relaxed_q15mulr_s_basic() {
    // Q15 multiply: (a * b + 0x4000) >> 15
    // With a=b=0x4000 (0.5 in Q15): (0x4000*0x4000 + 0x4000) >> 15 = 0x2000
    let v: i16 = 0x4000;
    let mut a = [0u8;16]; let mut b = [0u8;16];
    for i in 0..8 { a[i*2..i*2+2].copy_from_slice(&v.to_le_bytes()); b[i*2..i*2+2].copy_from_slice(&v.to_le_bytes()); }
    let r = as_v128(run(|c| { emit_v128(c, a); emit_v128(c, b); c.emit_op(Op::I16X8_RELAXED_Q15MULR_S, 0); }));
    let lane0 = i16::from_le_bytes([r[0], r[1]]);
    assert_eq!(lane0, 0x2000);
}

// ── i16x8.relaxed_dot_i8x16_i7x16_s ──────────────────────────────────────

#[test]
fn i16x8_relaxed_dot_i8x16_i7x16_s() {
    // dot: for each pair of adjacent i8 lanes, compute a[2i]*b[2i] + a[2i+1]*b[2i+1] → i16
    let mut a = [0i8;16]; let mut b = [0i8;16];
    a[0] = 3; a[1] = 4; b[0] = 3; b[1] = 4; // lane 0: 3*3 + 4*4 = 25
    let r = as_v128(run(|c| {
        emit_v128(c, a.map(|x| x as u8));
        emit_v128(c, b.map(|x| x as u8));
        c.emit_op(Op::I16X8_RELAXED_DOT_I8X16_I7X16_S, 0);
    }));
    let lane0 = i16::from_le_bytes([r[0], r[1]]);
    assert_eq!(lane0, 25);
}

// ── i32x4.relaxed_dot_i8x16_i7x16_add_s ──────────────────────────────────

#[test]
fn i32x4_relaxed_dot_i8x16_i7x16_add_s() {
    // dot_add: groups of 4 bytes: sum(a[4i+j]*b[4i+j]) for j=0..3 + accumulator lane
    // With all-1 inputs: each i32 lane = 1+1+1+1=4 (plus accumulator 0)
    let a: [u8;16] = [1u8;16];
    let b: [u8;16] = [1u8;16];
    let acc = mk_i32x4([0, 0, 0, 0]);
    let r = i32_lanes(&as_v128(run(|c| {
        emit_v128(c, a); emit_v128(c, b); emit_v128(c, acc);
        c.emit_op(Op::I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S, 0);
    })));
    assert_eq!(r, [4, 4, 4, 4]);
}
