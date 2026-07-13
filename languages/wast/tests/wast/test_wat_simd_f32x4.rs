//! SIMD f32x4 lane operations — 4 lanes of single-precision floats.
use crate::wat_exec;

wat_exec! {
    test_f32x4_splat => { r#"(func (export "_start")
        f32.const 2.5 f32x4.splat f32x4.extract_lane 1 call $log_f32)"#, "2.5" },
    test_f32x4_extract_lane => { r#"(func (export "_start")
        v128.const f32x4 1.5 2.5 3.5 4.5 f32x4.extract_lane 2 call $log_f32)"#, "3.5" },
    test_f32x4_replace_lane => { r#"(func (export "_start")
        v128.const f32x4 0 0 0 0 f32.const 9.5 f32x4.replace_lane 3
        f32x4.extract_lane 3 call $log_f32)"#, "9.5" },
    test_f32x4_add => { r#"(func (export "_start")
        v128.const f32x4 1.5 0 0 0 v128.const f32x4 2.5 0 0 0
        f32x4.add f32x4.extract_lane 0 call $log_f32)"#, "4.0" },
    test_f32x4_sub => { r#"(func (export "_start")
        v128.const f32x4 5.0 0 0 0 v128.const f32x4 1.5 0 0 0
        f32x4.sub f32x4.extract_lane 0 call $log_f32)"#, "3.5" },
    test_f32x4_mul => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 4.0 0 0 0
        f32x4.mul f32x4.extract_lane 0 call $log_f32)"#, "12.0" },
    test_f32x4_div => { r#"(func (export "_start")
        v128.const f32x4 9.0 0 0 0 v128.const f32x4 2.0 0 0 0
        f32x4.div f32x4.extract_lane 0 call $log_f32)"#, "4.5" },
    test_f32x4_min => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 8.0 0 0 0
        f32x4.min f32x4.extract_lane 0 call $log_f32)"#, "3.0" },
    test_f32x4_max => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 8.0 0 0 0
        f32x4.max f32x4.extract_lane 0 call $log_f32)"#, "8.0" },
    test_f32x4_pmin => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 8.0 0 0 0
        f32x4.pmin f32x4.extract_lane 0 call $log_f32)"#, "3.0" },
    test_f32x4_pmax => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 8.0 0 0 0
        f32x4.pmax f32x4.extract_lane 0 call $log_f32)"#, "8.0" },
    test_f32x4_abs => { r#"(func (export "_start")
        v128.const f32x4 -7.5 0 0 0 f32x4.abs f32x4.extract_lane 0 call $log_f32)"#, "7.5" },
    test_f32x4_neg => { r#"(func (export "_start")
        v128.const f32x4 7.5 0 0 0 f32x4.neg f32x4.extract_lane 0 call $log_f32)"#, "-7.5" },
    test_f32x4_sqrt => { r#"(func (export "_start")
        v128.const f32x4 16.0 0 0 0 f32x4.sqrt f32x4.extract_lane 0 call $log_f32)"#, "4.0" },
    test_f32x4_ceil => { r#"(func (export "_start")
        v128.const f32x4 2.1 0 0 0 f32x4.ceil f32x4.extract_lane 0 call $log_f32)"#, "3.0" },
    test_f32x4_floor => { r#"(func (export "_start")
        v128.const f32x4 2.9 0 0 0 f32x4.floor f32x4.extract_lane 0 call $log_f32)"#, "2.0" },
    test_f32x4_trunc => { r#"(func (export "_start")
        v128.const f32x4 -2.9 0 0 0 f32x4.trunc f32x4.extract_lane 0 call $log_f32)"#, "-2.0" },
    test_f32x4_nearest => { r#"(func (export "_start")
        v128.const f32x4 2.5 0 0 0 f32x4.nearest f32x4.extract_lane 0 call $log_f32)"#, "2.0" },
    test_f32x4_eq => { r#"(func (export "_start")
        v128.const f32x4 1.5 0 0 0 v128.const f32x4 1.5 0 0 0
        f32x4.eq i32x4.extract_lane 0 call $log)"#, "-1" },
    test_f32x4_lt => { r#"(func (export "_start")
        v128.const f32x4 1.0 0 0 0 v128.const f32x4 2.0 0 0 0
        f32x4.lt i32x4.extract_lane 0 call $log)"#, "-1" },
    test_f32x4_ge => { r#"(func (export "_start")
        v128.const f32x4 2.0 0 0 0 v128.const f32x4 2.0 0 0 0
        f32x4.ge i32x4.extract_lane 0 call $log)"#, "-1" },
    test_f32x4_convert_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 -5 0 0 0 f32x4.convert_i32x4_s f32x4.extract_lane 0 call $log_f32)"#, "-5.0" },
    test_f32x4_convert_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 -1 0 0 0 f32x4.convert_i32x4_u f32x4.extract_lane 0 call $log_f32)"#, "4294967296.0" },
    test_f32x4_demote_f64x2_zero => { r#"(func (export "_start")
        v128.const f64x2 3.5 0 f32x4.demote_f64x2_zero f32x4.extract_lane 0 call $log_f32)"#, "3.5" },

    // ── remaining float comparisons ──────────────────────────────────────────
    test_f32x4_ne => { r#"(func (export "_start")
        v128.const f32x4 1.0 0 0 0 v128.const f32x4 2.0 0 0 0
        f32x4.ne i32x4.extract_lane 0 call $log)"#, "-1" },
    test_f32x4_gt => { r#"(func (export "_start")
        v128.const f32x4 3.0 0 0 0 v128.const f32x4 2.0 0 0 0
        f32x4.gt i32x4.extract_lane 0 call $log)"#, "-1" },
    test_f32x4_le => { r#"(func (export "_start")
        v128.const f32x4 2.0 0 0 0 v128.const f32x4 2.0 0 0 0
        f32x4.le i32x4.extract_lane 0 call $log)"#, "-1" },
}
