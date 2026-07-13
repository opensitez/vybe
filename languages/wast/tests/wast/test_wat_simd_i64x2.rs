//! SIMD i64x2 lane operations — 2 lanes of 64-bit integers.
use crate::wat_exec;

wat_exec! {
    test_i64x2_splat => { r#"(func (export "_start")
        i64.const 5000000000 i64x2.splat i64x2.extract_lane 1 call $log_i64)"#, "5000000000" },
    test_i64x2_extract_lane => { r#"(func (export "_start")
        v128.const i64x2 111 222 i64x2.extract_lane 1 call $log_i64)"#, "222" },
    test_i64x2_replace_lane => { r#"(func (export "_start")
        v128.const i64x2 0 0 i64.const 999 i64x2.replace_lane 0
        i64x2.extract_lane 0 call $log_i64)"#, "999" },
    test_i64x2_add => { r#"(func (export "_start")
        v128.const i64x2 4000000000 0 v128.const i64x2 4000000000 0
        i64x2.add i64x2.extract_lane 0 call $log_i64)"#, "8000000000" },
    test_i64x2_sub => { r#"(func (export "_start")
        v128.const i64x2 10 0 v128.const i64x2 25 0
        i64x2.sub i64x2.extract_lane 0 call $log_i64)"#, "-15" },
    test_i64x2_mul => { r#"(func (export "_start")
        v128.const i64x2 100000 0 v128.const i64x2 100000 0
        i64x2.mul i64x2.extract_lane 0 call $log_i64)"#, "10000000000" },
    test_i64x2_neg => { r#"(func (export "_start")
        v128.const i64x2 42 0 i64x2.neg i64x2.extract_lane 0 call $log_i64)"#, "-42" },
    test_i64x2_abs => { r#"(func (export "_start")
        v128.const i64x2 -42 0 i64x2.abs i64x2.extract_lane 0 call $log_i64)"#, "42" },
    test_i64x2_shl => { r#"(func (export "_start")
        v128.const i64x2 1 0 i32.const 40 i64x2.shl i64x2.extract_lane 0 call $log_i64)"#, "1099511627776" },
    test_i64x2_shr_s => { r#"(func (export "_start")
        v128.const i64x2 -1024 0 i32.const 2 i64x2.shr_s i64x2.extract_lane 0 call $log_i64)"#, "-256" },
    test_i64x2_shr_u => { r#"(func (export "_start")
        v128.const i64x2 -1 0 i32.const 60 i64x2.shr_u i64x2.extract_lane 0 call $log_i64)"#, "15" },
    test_i64x2_eq => { r#"(func (export "_start")
        v128.const i64x2 7 0 v128.const i64x2 7 0
        i64x2.eq i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_i64x2_ne => { r#"(func (export "_start")
        v128.const i64x2 7 0 v128.const i64x2 8 0
        i64x2.ne i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_i64x2_lt_s => { r#"(func (export "_start")
        v128.const i64x2 -1 0 v128.const i64x2 0 0
        i64x2.lt_s i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_i64x2_all_true => { r#"(func (export "_start")
        v128.const i64x2 1 2 i64x2.all_true call $log)"#, "1" },
    test_i64x2_all_true_false => { r#"(func (export "_start")
        v128.const i64x2 1 0 i64x2.all_true call $log)"#, "0" },
    test_i64x2_bitmask => { r#"(func (export "_start")
        v128.const i64x2 -1 0 i64x2.bitmask call $log)"#, "1" },
    test_i64x2_extend_low_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 -5 0 0 0 i64x2.extend_low_i32x4_s i64x2.extract_lane 0 call $log_i64)"#, "-5" },
    test_i64x2_extend_low_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 i64x2.extend_low_i32x4_u i64x2.extract_lane 0 call $log_i64)"#, "4294967295" },
    test_i64x2_extmul_low_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 100000 0 0 0 v128.const i32x4 100000 0 0 0
        i64x2.extmul_low_i32x4_s i64x2.extract_lane 0 call $log_i64)"#, "10000000000" },

    // ── remaining signed comparisons + high widening (i64x2 has no unsigned) ──
    test_i64x2_gt_s => { r#"(func (export "_start")
        v128.const i64x2 5 0 v128.const i64x2 3 0
        i64x2.gt_s i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_i64x2_le_s => { r#"(func (export "_start")
        v128.const i64x2 -1 0 v128.const i64x2 0 0
        i64x2.le_s i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_i64x2_ge_s => { r#"(func (export "_start")
        v128.const i64x2 7 0 v128.const i64x2 7 0
        i64x2.ge_s i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_i64x2_extend_high_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 0 0 -5 0 i64x2.extend_high_i32x4_s i64x2.extract_lane 0 call $log_i64)"#, "-5" },
    test_i64x2_extend_high_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 0 0 -1 0 i64x2.extend_high_i32x4_u i64x2.extract_lane 0 call $log_i64)"#, "4294967295" },
    test_i64x2_extmul_high_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 0 0 100000 0 v128.const i32x4 0 0 100000 0
        i64x2.extmul_high_i32x4_s i64x2.extract_lane 0 call $log_i64)"#, "10000000000" },
    test_i64x2_extmul_low_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 v128.const i32x4 2 0 0 0
        i64x2.extmul_low_i32x4_u i64x2.extract_lane 0 call $log_i64)"#, "8589934590" },
    test_i64x2_extmul_high_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 0 0 -1 0 v128.const i32x4 0 0 2 0
        i64x2.extmul_high_i32x4_u i64x2.extract_lane 0 call $log_i64)"#, "8589934590" },
}
