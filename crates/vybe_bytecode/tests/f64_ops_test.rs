//! Tests for all f64 instructions from the WASM spec (§5.3 numeric).
//! Covers: const, comparisons (eq/ne/lt/gt/le/ge),
//!         unary (abs/neg/ceil/floor/trunc/nearest/sqrt),
//!         binary (add/sub/mul/div/min/max/copysign).

use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::value::Value;

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn push(c: &mut Chunk, v: f64) { let k = c.add_constant(Value::F64(v)); c.emit_op_u16(Op::CONST, k, 0); }

// ── f64.const ────────────────────────────────────────────────────────────

#[test] fn f64_const_zero()     { assert_eq!(run(|c| push(c,0.0)).as_f64(), 0.0); }
#[test] fn f64_const_pi()       { let r = run(|c| push(c, std::f64::consts::PI)); assert!((r.as_f64() - std::f64::consts::PI).abs() < 1e-15); }
#[test] fn f64_const_negative() { assert_eq!(run(|c| push(c,-1.5)).as_f64(), -1.5); }
#[test] fn f64_const_inf()      { assert!(run(|c| push(c, f64::INFINITY)).as_f64().is_infinite()); }
#[test] fn f64_const_nan()      { assert!(run(|c| push(c, f64::NAN)).as_f64().is_nan()); }

// ── f64 comparisons ───────────────────────────────────────────────────────

#[test] fn f64_eq_true()    { assert_eq!(run(|c| { push(c,1.0); push(c,1.0); c.emit_op(Op::F64_EQ,0); }).as_i32(), 1); }
#[test] fn f64_eq_false()   { assert_eq!(run(|c| { push(c,1.0); push(c,2.0); c.emit_op(Op::F64_EQ,0); }).as_i32(), 0); }
#[test] fn f64_ne_true()    { assert_eq!(run(|c| { push(c,1.0); push(c,2.0); c.emit_op(Op::F64_NE,0); }).as_i32(), 1); }
#[test] fn f64_ne_false()   { assert_eq!(run(|c| { push(c,1.0); push(c,1.0); c.emit_op(Op::F64_NE,0); }).as_i32(), 0); }
#[test] fn f64_lt_true()    { assert_eq!(run(|c| { push(c,1.0); push(c,2.0); c.emit_op(Op::F64_LT,0); }).as_i32(), 1); }
#[test] fn f64_lt_false()   { assert_eq!(run(|c| { push(c,2.0); push(c,1.0); c.emit_op(Op::F64_LT,0); }).as_i32(), 0); }
#[test] fn f64_gt_true()    { assert_eq!(run(|c| { push(c,3.0); push(c,1.0); c.emit_op(Op::F64_GT,0); }).as_i32(), 1); }
#[test] fn f64_gt_false()   { assert_eq!(run(|c| { push(c,1.0); push(c,3.0); c.emit_op(Op::F64_GT,0); }).as_i32(), 0); }
#[test] fn f64_le_equal()   { assert_eq!(run(|c| { push(c,2.0); push(c,2.0); c.emit_op(Op::F64_LE,0); }).as_i32(), 1); }
#[test] fn f64_le_less()    { assert_eq!(run(|c| { push(c,1.0); push(c,2.0); c.emit_op(Op::F64_LE,0); }).as_i32(), 1); }
#[test] fn f64_ge_equal()   { assert_eq!(run(|c| { push(c,2.0); push(c,2.0); c.emit_op(Op::F64_GE,0); }).as_i32(), 1); }
#[test] fn f64_ge_greater() { assert_eq!(run(|c| { push(c,3.0); push(c,2.0); c.emit_op(Op::F64_GE,0); }).as_i32(), 1); }
#[test] fn f64_eq_nan_false(){ assert_eq!(run(|c| { push(c,f64::NAN); push(c,f64::NAN); c.emit_op(Op::F64_EQ,0); }).as_i32(), 0); }
#[test] fn f64_ne_nan_true() { assert_eq!(run(|c| { push(c,f64::NAN); push(c,1.0); c.emit_op(Op::F64_NE,0); }).as_i32(), 1); }
#[test] fn f64_lt_nan_false(){ assert_eq!(run(|c| { push(c,f64::NAN); push(c,1.0); c.emit_op(Op::F64_LT,0); }).as_i32(), 0); }

// ── f64 unary ─────────────────────────────────────────────────────────────

#[test] fn f64_abs_positive() { assert_eq!(run(|c| { push(c, 3.0); c.emit_op(Op::F64_ABS,0); }).as_f64(), 3.0); }
#[test] fn f64_abs_negative() { assert_eq!(run(|c| { push(c,-3.0); c.emit_op(Op::F64_ABS,0); }).as_f64(), 3.0); }
#[test] fn f64_neg()          { assert_eq!(run(|c| { push(c, 3.0); c.emit_op(Op::F64_NEG,0); }).as_f64(), -3.0); }
#[test] fn f64_ceil()         { assert_eq!(run(|c| { push(c, 1.2); c.emit_op(Op::F64_CEIL,0); }).as_f64(), 2.0); }
#[test] fn f64_floor()        { assert_eq!(run(|c| { push(c, 1.8); c.emit_op(Op::F64_FLOOR,0); }).as_f64(), 1.0); }
#[test] fn f64_trunc_pos()    { assert_eq!(run(|c| { push(c, 1.9); c.emit_op(Op::F64_TRUNC,0); }).as_f64(), 1.0); }
#[test] fn f64_trunc_neg()    { assert_eq!(run(|c| { push(c,-1.9); c.emit_op(Op::F64_TRUNC,0); }).as_f64(), -1.0); }
#[test] fn f64_nearest_half() { assert_eq!(run(|c| { push(c, 0.5); c.emit_op(Op::F64_NEAREST,0); }).as_f64(), 0.0); }
#[test] fn f64_sqrt()         { assert_eq!(run(|c| { push(c, 9.0); c.emit_op(Op::F64_SQRT,0); }).as_f64(), 3.0); }
#[test] fn f64_sqrt_nan()     { assert!(run(|c| { push(c,-1.0); c.emit_op(Op::F64_SQRT,0); }).as_f64().is_nan()); }

// ── f64 binary ────────────────────────────────────────────────────────────

#[test] fn f64_add() { assert_eq!(run(|c| { push(c,1.5); push(c,2.5); c.emit_op(Op::F64_ADD,0); }).as_f64(), 4.0); }
#[test] fn f64_sub() { assert_eq!(run(|c| { push(c,5.0); push(c,3.0); c.emit_op(Op::F64_SUB,0); }).as_f64(), 2.0); }
#[test] fn f64_mul() { assert_eq!(run(|c| { push(c,3.0); push(c,4.0); c.emit_op(Op::F64_MUL,0); }).as_f64(), 12.0); }
#[test] fn f64_div() { assert_eq!(run(|c| { push(c,7.0); push(c,2.0); c.emit_op(Op::F64_DIV,0); }).as_f64(), 3.5); }
#[test] fn f64_min() { assert_eq!(run(|c| { push(c,1.0); push(c,2.0); c.emit_op(Op::F64_MIN,0); }).as_f64(), 1.0); }
#[test] fn f64_max() { assert_eq!(run(|c| { push(c,1.0); push(c,2.0); c.emit_op(Op::F64_MAX,0); }).as_f64(), 2.0); }
#[test] fn f64_copysign_positive() {
    assert_eq!(run(|c| { push(c,-3.0); push(c,1.0); c.emit_op(Op::F64_COPYSIGN,0); }).as_f64(), 3.0);
}
#[test] fn f64_copysign_negative() {
    assert_eq!(run(|c| { push(c,3.0); push(c,-1.0); c.emit_op(Op::F64_COPYSIGN,0); }).as_f64(), -3.0);
}
#[test] fn f64_div_by_zero_infinity() {
    assert!(run(|c| { push(c,1.0); push(c,0.0); c.emit_op(Op::F64_DIV,0); }).as_f64().is_infinite());
}

// ── Spec-required edge cases ──────────────────────────────────────────────

#[test] fn f64_nearest_one_point_five()  { assert_eq!(run(|c| { push(c,1.5); c.emit_op(Op::F64_NEAREST,0); }).as_f64(), 2.0); }
#[test] fn f64_nearest_two_point_five()  { assert_eq!(run(|c| { push(c,2.5); c.emit_op(Op::F64_NEAREST,0); }).as_f64(), 2.0); }
#[test] fn f64_nearest_neg_half()        { assert_eq!(run(|c| { push(c,-0.5); c.emit_op(Op::F64_NEAREST,0); }).as_f64(), 0.0); }
#[test] fn f64_min_nan_propagates()      { assert!(run(|c| { push(c,f64::NAN); push(c,5.0); c.emit_op(Op::F64_MIN,0); }).as_f64().is_nan()); }
#[test] fn f64_max_nan_propagates()      { assert!(run(|c| { push(c,f64::NAN); push(c,5.0); c.emit_op(Op::F64_MAX,0); }).as_f64().is_nan()); }
#[test] fn f64_min_nan_second()          { assert!(run(|c| { push(c,5.0); push(c,f64::NAN); c.emit_op(Op::F64_MIN,0); }).as_f64().is_nan()); }
#[test] fn f64_min_neg_zero()            { assert!(run(|c| { push(c,-0.0); push(c,0.0); c.emit_op(Op::F64_MIN,0); }).as_f64().is_sign_negative()); }
#[test] fn f64_max_pos_zero()            { assert!(run(|c| { push(c,-0.0); push(c,0.0); c.emit_op(Op::F64_MAX,0); }).as_f64().is_sign_positive()); }
