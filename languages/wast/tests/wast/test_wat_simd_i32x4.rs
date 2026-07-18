//! SIMD i32x4 lane operations — 4 lanes of 32-bit integers.
//! Expected values from the WebAssembly SIMD spec.
use crate::wat_exec;

wat_exec! {
    test_i32x4_splat => { r#"(func (export "_start")
        i32.const 123456 i32x4.splat i32x4.extract_lane 2 call $log)"#, "123456" },
    test_i32x4_extract_lane => { r#"(func (export "_start")
        v128.const i32x4 10 20 30 40 i32x4.extract_lane 3 call $log)"#, "40" },
    test_i32x4_replace_lane => { r#"(func (export "_start")
        v128.const i32x4 10 20 30 40 i32.const 99 i32x4.replace_lane 1
        i32x4.extract_lane 1 call $log)"#, "99" },

    test_i32x4_add => { r#"(func (export "_start")
        v128.const i32x4 100 0 0 0 v128.const i32x4 23 0 0 0
        i32x4.add i32x4.extract_lane 0 call $log)"#, "123" },
    test_i32x4_add_wraps => { r#"(func (export "_start")
        v128.const i32x4 2147483647 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.add i32x4.extract_lane 0 call $log)"#, "-2147483648" },
    test_i32x4_sub => { r#"(func (export "_start")
        v128.const i32x4 5 0 0 0 v128.const i32x4 8 0 0 0
        i32x4.sub i32x4.extract_lane 0 call $log)"#, "-3" },
    test_i32x4_mul => { r#"(func (export "_start")
        v128.const i32x4 6 0 0 0 v128.const i32x4 7 0 0 0
        i32x4.mul i32x4.extract_lane 0 call $log)"#, "42" },
    test_i32x4_neg => { r#"(func (export "_start")
        v128.const i32x4 55 0 0 0 i32x4.neg i32x4.extract_lane 0 call $log)"#, "-55" },
    test_i32x4_abs => { r#"(func (export "_start")
        v128.const i32x4 -55 0 0 0 i32x4.abs i32x4.extract_lane 0 call $log)"#, "55" },

    test_i32x4_min_s => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.min_s i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_min_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.min_u i32x4.extract_lane 0 call $log)"#, "1" },
    test_i32x4_max_s => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.max_s i32x4.extract_lane 0 call $log)"#, "1" },
    test_i32x4_max_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.max_u i32x4.extract_lane 0 call $log)"#, "-1" },

    test_i32x4_shl => { r#"(func (export "_start")
        v128.const i32x4 1 0 0 0 i32.const 10 i32x4.shl i32x4.extract_lane 0 call $log)"#, "1024" },
    test_i32x4_shr_s => { r#"(func (export "_start")
        v128.const i32x4 -256 0 0 0 i32.const 2 i32x4.shr_s i32x4.extract_lane 0 call $log)"#, "-64" },
    test_i32x4_shr_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 i32.const 28 i32x4.shr_u i32x4.extract_lane 0 call $log)"#, "15" },

    test_i32x4_eq => { r#"(func (export "_start")
        v128.const i32x4 7 0 0 0 v128.const i32x4 7 0 0 0
        i32x4.eq i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_ne => { r#"(func (export "_start")
        v128.const i32x4 7 0 0 0 v128.const i32x4 8 0 0 0
        i32x4.ne i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_lt_s => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 0 0 0 0
        i32x4.lt_s i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_lt_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 0 0 0 0
        i32x4.lt_u i32x4.extract_lane 0 call $log)"#, "0" },
    test_i32x4_ge_s => { r#"(func (export "_start")
        v128.const i32x4 5 0 0 0 v128.const i32x4 5 0 0 0
        i32x4.ge_s i32x4.extract_lane 0 call $log)"#, "-1" },

    test_i32x4_all_true => { r#"(func (export "_start")
        v128.const i32x4 1 2 3 4 i32x4.all_true call $log)"#, "1" },
    test_i32x4_all_true_false => { r#"(func (export "_start")
        v128.const i32x4 1 0 3 4 i32x4.all_true call $log)"#, "0" },
    test_i32x4_bitmask => { r#"(func (export "_start")
        v128.const i32x4 -1 0 -1 0 i32x4.bitmask call $log)"#, "5" },

    test_i32x4_dot_i16x8_s => { r#"(func (export "_start")
        v128.const i16x8 2 3 0 0 0 0 0 0 v128.const i16x8 4 5 0 0 0 0 0 0
        i32x4.dot_i16x8_s i32x4.extract_lane 0 call $log)"#, "23" },
    test_i32x4_extend_low_i16x8_s => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0
        i32x4.extend_low_i16x8_s i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_extend_low_i16x8_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0
        i32x4.extend_low_i16x8_u i32x4.extract_lane 0 call $log)"#, "65535" },
    test_i32x4_extmul_low_i16x8_s => { r#"(func (export "_start")
        v128.const i16x8 1000 0 0 0 0 0 0 0 v128.const i16x8 1000 0 0 0 0 0 0 0
        i32x4.extmul_low_i16x8_s i32x4.extract_lane 0 call $log)"#, "1000000" },
    test_i32x4_extadd_pairwise_i16x8_s => { r#"(func (export "_start")
        v128.const i16x8 100 200 0 0 0 0 0 0
        i32x4.extadd_pairwise_i16x8_s i32x4.extract_lane 0 call $log)"#, "300" },

    test_i32x4_trunc_sat_f32x4_s => { r#"(func (export "_start")
        v128.const f32x4 3.9 0 0 0 i32x4.trunc_sat_f32x4_s i32x4.extract_lane 0 call $log)"#, "3" },
    test_i32x4_trunc_sat_f32x4_s_clamps => { r#"(func (export "_start")
        v128.const f32x4 3e10 0 0 0 i32x4.trunc_sat_f32x4_s i32x4.extract_lane 0 call $log)"#, "2147483647" },
    test_i32x4_trunc_sat_f32x4_u_negative_zero => { r#"(func (export "_start")
        v128.const f32x4 -1.0 0 0 0 i32x4.trunc_sat_f32x4_u i32x4.extract_lane 0 call $log)"#, "0" },

    // ── remaining comparisons + widening variants ────────────────────────────
    test_i32x4_gt_s => { r#"(func (export "_start")
        v128.const i32x4 5 0 0 0 v128.const i32x4 3 0 0 0
        i32x4.gt_s i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_gt_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.gt_u i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_le_s => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 0 0 0 0
        i32x4.le_s i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_le_u => { r#"(func (export "_start")
        v128.const i32x4 1 0 0 0 v128.const i32x4 -1 0 0 0
        i32x4.le_u i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_ge_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 1 0 0 0
        i32x4.ge_u i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_extend_high_i16x8_s => { r#"(func (export "_start")
        v128.const i16x8 0 0 0 0 0xFFFF 0 0 0
        i32x4.extend_high_i16x8_s i32x4.extract_lane 0 call $log)"#, "-1" },
    test_i32x4_extend_high_i16x8_u => { r#"(func (export "_start")
        v128.const i16x8 0 0 0 0 0xFFFF 0 0 0
        i32x4.extend_high_i16x8_u i32x4.extract_lane 0 call $log)"#, "65535" },
    test_i32x4_extmul_high_i16x8_s => { r#"(func (export "_start")
        v128.const i16x8 0 0 0 0 1000 0 0 0 v128.const i16x8 0 0 0 0 1000 0 0 0
        i32x4.extmul_high_i16x8_s i32x4.extract_lane 0 call $log)"#, "1000000" },
    test_i32x4_extmul_low_i16x8_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 2 0 0 0 0 0 0 0
        i32x4.extmul_low_i16x8_u i32x4.extract_lane 0 call $log)"#, "131070" },
    test_i32x4_extadd_pairwise_i16x8_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0xFFFF 0 0 0 0 0 0
        i32x4.extadd_pairwise_i16x8_u i32x4.extract_lane 0 call $log)"#, "131070" },
    test_i32x4_extmul_high_i16x8_u => { r#"(func (export "_start")
        v128.const i16x8 0 0 0 0 0xFFFF 0 0 0 v128.const i16x8 0 0 0 0 2 0 0 0
        i32x4.extmul_high_i16x8_u i32x4.extract_lane 0 call $log)"#, "131070" },
    test_i32x4_trunc_sat_f64x2_s_zero => { r#"(func (export "_start")
        v128.const f64x2 3.9 0 i32x4.trunc_sat_f64x2_s_zero i32x4.extract_lane 0 call $log)"#, "3" },
    test_i32x4_trunc_sat_f64x2_u_zero => { r#"(func (export "_start")
        v128.const f64x2 3.9 0 i32x4.trunc_sat_f64x2_u_zero i32x4.extract_lane 0 call $log)"#, "3" },

    // ── Spec edge cases: trunc_sat NaN→0, high-lane zeroing, signed dot ───────
    // Every `trunc_sat` maps NaN to 0 (not the saturation bound).
    test_i32x4_trunc_sat_f32x4_s_nan_is_zero => { r#"(func (export "_start")
        v128.const f32x4 nan 0 0 0 i32x4.trunc_sat_f32x4_s i32x4.extract_lane 0 call $log)"#, "0" },
    test_i32x4_trunc_sat_f32x4_s_neg_clamps => { r#"(func (export "_start")
        v128.const f32x4 -3e10 0 0 0 i32x4.trunc_sat_f32x4_s i32x4.extract_lane 0 call $log)"#, "-2147483648" },
    // The `_zero` variants zero the upper two i32 lanes (from the absent f64 lanes).
    test_i32x4_trunc_sat_f64x2_s_zero_upper_lane_is_zero => { r#"(func (export "_start")
        v128.const f64x2 3.9 7.9 i32x4.trunc_sat_f64x2_s_zero i32x4.extract_lane 2 call $log)"#, "0" },
    // dot with a negative operand: (-2)*4 + 3*5 = -8 + 15 = 7.
    test_i32x4_dot_i16x8_s_signed => { r#"(func (export "_start")
        v128.const i16x8 -2 3 0 0 0 0 0 0 v128.const i16x8 4 5 0 0 0 0 0 0
        i32x4.dot_i16x8_s i32x4.extract_lane 0 call $log)"#, "7" },
}
