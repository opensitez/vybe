//! Relaxed SIMD proposal — fused/relaxed lane operations. Results are given
//! for the standard (non-relaxed) rounding, which conforming engines produce
//! for these inputs.
use crate::wat_exec;

wat_exec! {
    test_i8x16_relaxed_swizzle => { r#"(func (export "_start")
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 5 15 25 35
        v128.const i8x16 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.relaxed_swizzle i8x16.extract_lane_u 0 call $log)"#, "40" },
    test_f32x4_relaxed_madd => { r#"(func (export "_start")
        v128.const f32x4 2.0 0 0 0 v128.const f32x4 3.0 0 0 0 v128.const f32x4 4.0 0 0 0
        f32x4.relaxed_madd f32x4.extract_lane 0 call $log_f32)"#, "10.0" },
    test_f32x4_relaxed_nmadd => { r#"(func (export "_start")
        v128.const f32x4 2.0 0 0 0 v128.const f32x4 3.0 0 0 0 v128.const f32x4 4.0 0 0 0
        f32x4.relaxed_nmadd f32x4.extract_lane 0 call $log_f32)"#, "-2.0" },
    test_f32x4_relaxed_min => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 8.0 0 0 0
        f32x4.relaxed_min f32x4.extract_lane 0 call $log_f32)"#, "3.0" },
    test_f32x4_relaxed_max => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 8.0 0 0 0
        f32x4.relaxed_max f32x4.extract_lane 0 call $log_f32)"#, "8.0" },
    test_i32x4_relaxed_trunc_f32x4_s => { r#"(func (export "_start")
        v128.const f32x4 3.9 0 0 0 i32x4.relaxed_trunc_f32x4_s i32x4.extract_lane 0 call $log)"#, "3" },
    test_i16x8_relaxed_q15mulr_s => { r#"(func (export "_start")
        v128.const i16x8 16384 0 0 0 0 0 0 0 v128.const i16x8 16384 0 0 0 0 0 0 0
        i16x8.relaxed_q15mulr_s i16x8.extract_lane_s 0 call $log)"#, "8192" },
    test_i8x16_relaxed_laneselect => { r#"(func (export "_start")
        v128.const i8x16 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1
        v128.const i8x16 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2
        v128.const i8x16 0xFF 0 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.relaxed_laneselect i8x16.extract_lane_u 0 call $log)"#, "1" },

    // ── remaining relaxed ops (f64x2 fma/min/max, laneselects, trunc) ────────
    test_f64x2_relaxed_madd => { r#"(func (export "_start")
        v128.const f64x2 2.0 0 v128.const f64x2 3.0 0 v128.const f64x2 4.0 0
        f64x2.relaxed_madd f64x2.extract_lane 0 call $log_f64)"#, "10.0" },
    test_f64x2_relaxed_nmadd => { r#"(func (export "_start")
        v128.const f64x2 2.0 0 v128.const f64x2 3.0 0 v128.const f64x2 4.0 0
        f64x2.relaxed_nmadd f64x2.extract_lane 0 call $log_f64)"#, "-2.0" },
    test_f64x2_relaxed_min => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.relaxed_min f64x2.extract_lane 0 call $log_f64)"#, "3.0" },
    test_f64x2_relaxed_max => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.relaxed_max f64x2.extract_lane 0 call $log_f64)"#, "8.0" },
    test_i16x8_relaxed_laneselect => { r#"(func (export "_start")
        v128.const i16x8 1 1 1 1 1 1 1 1 v128.const i16x8 2 2 2 2 2 2 2 2
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0
        i16x8.relaxed_laneselect i16x8.extract_lane_s 0 call $log)"#, "1" },
    test_i32x4_relaxed_laneselect => { r#"(func (export "_start")
        v128.const i32x4 1 1 1 1 v128.const i32x4 2 2 2 2 v128.const i32x4 -1 0 0 0
        i32x4.relaxed_laneselect i32x4.extract_lane 0 call $log)"#, "1" },
    test_i64x2_relaxed_laneselect => { r#"(func (export "_start")
        v128.const i64x2 1 1 v128.const i64x2 2 2 v128.const i64x2 -1 0
        i64x2.relaxed_laneselect i64x2.extract_lane 0 call $log_i64)"#, "1" },
    test_i32x4_relaxed_trunc_f32x4_u => { r#"(func (export "_start")
        v128.const f32x4 5.9 0 0 0 i32x4.relaxed_trunc_f32x4_u i32x4.extract_lane 0 call $log)"#, "5" },
    test_i32x4_relaxed_trunc_f64x2_s_zero => { r#"(func (export "_start")
        v128.const f64x2 7.9 0 i32x4.relaxed_trunc_f64x2_s_zero i32x4.extract_lane 0 call $log)"#, "7" },
    test_i32x4_relaxed_trunc_f64x2_u_zero => { r#"(func (export "_start")
        v128.const f64x2 7.9 0 i32x4.relaxed_trunc_f64x2_u_zero i32x4.extract_lane 0 call $log)"#, "7" },

    // ── relaxed integer dot products (i8×i7 → i16 / i32 accumulate) ───────────
    // lane0 = a0*b0 + a1*b1 = 2*4 + 3*5 = 23 (b masked to 7 bits, here in range).
    test_i16x8_relaxed_dot_i8x16_i7x16_s => { r#"(func (export "_start")
        v128.const i8x16 2 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 4 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.relaxed_dot_i8x16_i7x16_s i16x8.extract_lane_s 0 call $log)"#, "23" },
    // lane0 = c0 + (a0*b0 + a1*b1 + a2*b2 + a3*b3) = 100 + (1+2+3+4)*2 = 120.
    test_i32x4_relaxed_dot_i8x16_i7x16_add_s => { r#"(func (export "_start")
        v128.const i8x16 1 2 3 4 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 2 2 2 2 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i32x4 100 0 0 0
        i32x4.relaxed_dot_i8x16_i7x16_add_s i32x4.extract_lane 0 call $log)"#, "120" },
}
