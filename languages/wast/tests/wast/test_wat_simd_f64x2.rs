//! SIMD f64x2 lane operations — 2 lanes of double-precision floats.
use crate::wat_exec;

wat_exec! {
    test_f64x2_splat => { r#"(func (export "_start")
        f64.const 6.25 f64x2.splat f64x2.extract_lane 1 call $log_f64)"#, "6.25" },
    test_f64x2_extract_lane => { r#"(func (export "_start")
        v128.const f64x2 1.5 9.5 f64x2.extract_lane 1 call $log_f64)"#, "9.5" },
    test_f64x2_replace_lane => { r#"(func (export "_start")
        v128.const f64x2 0 0 f64.const 7.75 f64x2.replace_lane 0
        f64x2.extract_lane 0 call $log_f64)"#, "7.75" },
    test_f64x2_add => { r#"(func (export "_start")
        v128.const f64x2 1.25 0 v128.const f64x2 2.75 0
        f64x2.add f64x2.extract_lane 0 call $log_f64)"#, "4.0" },
    test_f64x2_sub => { r#"(func (export "_start")
        v128.const f64x2 5.0 0 v128.const f64x2 1.25 0
        f64x2.sub f64x2.extract_lane 0 call $log_f64)"#, "3.75" },
    test_f64x2_mul => { r#"(func (export "_start")
        v128.const f64x2 2.5 0 v128.const f64x2 4.0 0
        f64x2.mul f64x2.extract_lane 0 call $log_f64)"#, "10.0" },
    test_f64x2_div => { r#"(func (export "_start")
        v128.const f64x2 7.0 0 v128.const f64x2 2.0 0
        f64x2.div f64x2.extract_lane 0 call $log_f64)"#, "3.5" },
    test_f64x2_min => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.min f64x2.extract_lane 0 call $log_f64)"#, "3.0" },
    test_f64x2_max => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.max f64x2.extract_lane 0 call $log_f64)"#, "8.0" },
    test_f64x2_abs => { r#"(func (export "_start")
        v128.const f64x2 -9.5 0 f64x2.abs f64x2.extract_lane 0 call $log_f64)"#, "9.5" },
    test_f64x2_neg => { r#"(func (export "_start")
        v128.const f64x2 9.5 0 f64x2.neg f64x2.extract_lane 0 call $log_f64)"#, "-9.5" },
    test_f64x2_sqrt => { r#"(func (export "_start")
        v128.const f64x2 25.0 0 f64x2.sqrt f64x2.extract_lane 0 call $log_f64)"#, "5.0" },
    test_f64x2_ceil => { r#"(func (export "_start")
        v128.const f64x2 2.1 0 f64x2.ceil f64x2.extract_lane 0 call $log_f64)"#, "3.0" },
    test_f64x2_floor => { r#"(func (export "_start")
        v128.const f64x2 2.9 0 f64x2.floor f64x2.extract_lane 0 call $log_f64)"#, "2.0" },
    test_f64x2_nearest => { r#"(func (export "_start")
        v128.const f64x2 3.5 0 f64x2.nearest f64x2.extract_lane 0 call $log_f64)"#, "4.0" },
    test_f64x2_eq => { r#"(func (export "_start")
        v128.const f64x2 1.5 0 v128.const f64x2 1.5 0
        f64x2.eq i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_f64x2_lt => { r#"(func (export "_start")
        v128.const f64x2 1.0 0 v128.const f64x2 2.0 0
        f64x2.lt i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_f64x2_convert_low_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 -7 0 0 0 f64x2.convert_low_i32x4_s f64x2.extract_lane 0 call $log_f64)"#, "-7.0" },
    test_f64x2_promote_low_f32x4 => { r#"(func (export "_start")
        v128.const f32x4 2.5 0 0 0 f64x2.promote_low_f32x4 f64x2.extract_lane 0 call $log_f64)"#, "2.5" },

    // ── remaining comparisons, pmin/pmax, trunc, convert ─────────────────────
    test_f64x2_ne => { r#"(func (export "_start")
        v128.const f64x2 1.0 0 v128.const f64x2 2.0 0
        f64x2.ne i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_f64x2_gt => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 2.0 0
        f64x2.gt i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_f64x2_le => { r#"(func (export "_start")
        v128.const f64x2 2.0 0 v128.const f64x2 2.0 0
        f64x2.le i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_f64x2_ge => { r#"(func (export "_start")
        v128.const f64x2 5.0 0 v128.const f64x2 5.0 0
        f64x2.ge i64x2.extract_lane 0 call $log_i64)"#, "-1" },
    test_f64x2_pmin => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.pmin f64x2.extract_lane 0 call $log_f64)"#, "3.0" },
    test_f64x2_pmax => { r#"(func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.pmax f64x2.extract_lane 0 call $log_f64)"#, "8.0" },
    test_f64x2_trunc => { r#"(func (export "_start")
        v128.const f64x2 -2.9 0 f64x2.trunc f64x2.extract_lane 0 call $log_f64)"#, "-2.0" },
    test_f64x2_convert_low_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 f64x2.convert_low_i32x4_u f64x2.extract_lane 0 call $log_f64)"#, "4294967295.0" },
}
