//! SIMD i16x8 lane operations — 8 lanes of 16-bit integers.
//! Expected values from the WebAssembly SIMD spec: 16-bit wrapping, signed/
//! unsigned saturation and extraction, extend/extmul from i8x16.
use crate::wat_exec;

wat_exec! {
    test_i16x8_splat => { r#"(func (export "_start")
        i32.const 1000 i16x8.splat i16x8.extract_lane_s 3 call $log)"#, "1000" },
    test_i16x8_splat_truncates_to_16 => { r#"(func (export "_start")
        i32.const 0x1FFFF i16x8.splat i16x8.extract_lane_u 0 call $log)"#, "65535" },
    test_i16x8_extract_lane_s => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_extract_lane_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 i16x8.extract_lane_u 0 call $log)"#, "65535" },
    test_i16x8_replace_lane => { r#"(func (export "_start")
        v128.const i16x8 0 0 0 0 0 0 0 0 i32.const 777 i16x8.replace_lane 4
        i16x8.extract_lane_s 4 call $log)"#, "777" },

    test_i16x8_add_wraps => { r#"(func (export "_start")
        v128.const i16x8 32767 0 0 0 0 0 0 0 v128.const i16x8 1 0 0 0 0 0 0 0
        i16x8.add i16x8.extract_lane_s 0 call $log)"#, "-32768" },
    test_i16x8_sub => { r#"(func (export "_start")
        v128.const i16x8 10 0 0 0 0 0 0 0 v128.const i16x8 25 0 0 0 0 0 0 0
        i16x8.sub i16x8.extract_lane_s 0 call $log)"#, "-15" },
    test_i16x8_mul => { r#"(func (export "_start")
        v128.const i16x8 300 0 0 0 0 0 0 0 v128.const i16x8 300 0 0 0 0 0 0 0
        i16x8.mul i16x8.extract_lane_s 0 call $log)"#, "24464" },
    test_i16x8_neg => { r#"(func (export "_start")
        v128.const i16x8 500 0 0 0 0 0 0 0 i16x8.neg i16x8.extract_lane_s 0 call $log)"#, "-500" },
    test_i16x8_abs => { r#"(func (export "_start")
        v128.const i16x8 0xFF00 0 0 0 0 0 0 0 i16x8.abs i16x8.extract_lane_s 0 call $log)"#, "256" },

    test_i16x8_add_sat_s => { r#"(func (export "_start")
        v128.const i16x8 32767 0 0 0 0 0 0 0 v128.const i16x8 100 0 0 0 0 0 0 0
        i16x8.add_sat_s i16x8.extract_lane_s 0 call $log)"#, "32767" },
    test_i16x8_add_sat_u => { r#"(func (export "_start")
        v128.const i16x8 65535 0 0 0 0 0 0 0 v128.const i16x8 1 0 0 0 0 0 0 0
        i16x8.add_sat_u i16x8.extract_lane_u 0 call $log)"#, "65535" },
    test_i16x8_sub_sat_s => { r#"(func (export "_start")
        v128.const i16x8 0x8000 0 0 0 0 0 0 0 v128.const i16x8 1 0 0 0 0 0 0 0
        i16x8.sub_sat_s i16x8.extract_lane_s 0 call $log)"#, "-32768" },
    test_i16x8_sub_sat_u => { r#"(func (export "_start")
        v128.const i16x8 5 0 0 0 0 0 0 0 v128.const i16x8 10 0 0 0 0 0 0 0
        i16x8.sub_sat_u i16x8.extract_lane_u 0 call $log)"#, "0" },

    test_i16x8_min_s => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 5 0 0 0 0 0 0 0
        i16x8.min_s i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_min_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 5 0 0 0 0 0 0 0
        i16x8.min_u i16x8.extract_lane_u 0 call $log)"#, "5" },
    test_i16x8_max_s => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 5 0 0 0 0 0 0 0
        i16x8.max_s i16x8.extract_lane_s 0 call $log)"#, "5" },
    test_i16x8_max_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 5 0 0 0 0 0 0 0
        i16x8.max_u i16x8.extract_lane_u 0 call $log)"#, "65535" },

    test_i16x8_shl => { r#"(func (export "_start")
        v128.const i16x8 1 0 0 0 0 0 0 0 i32.const 4 i16x8.shl i16x8.extract_lane_s 0 call $log)"#, "16" },
    test_i16x8_shr_s => { r#"(func (export "_start")
        v128.const i16x8 0x8000 0 0 0 0 0 0 0 i32.const 1 i16x8.shr_s
        i16x8.extract_lane_s 0 call $log)"#, "-16384" },
    test_i16x8_shr_u => { r#"(func (export "_start")
        v128.const i16x8 0x8000 0 0 0 0 0 0 0 i32.const 1 i16x8.shr_u
        i16x8.extract_lane_u 0 call $log)"#, "16384" },
    test_i16x8_avgr_u => { r#"(func (export "_start")
        v128.const i16x8 5 0 0 0 0 0 0 0 v128.const i16x8 8 0 0 0 0 0 0 0
        i16x8.avgr_u i16x8.extract_lane_u 0 call $log)"#, "7" },

    test_i16x8_eq_true => { r#"(func (export "_start")
        v128.const i16x8 42 0 0 0 0 0 0 0 v128.const i16x8 42 0 0 0 0 0 0 0
        i16x8.eq i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_lt_s => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 0 0 0 0 0 0 0 0
        i16x8.lt_s i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_gt_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 1 0 0 0 0 0 0 0
        i16x8.gt_u i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_all_true => { r#"(func (export "_start")
        v128.const i16x8 1 2 3 4 5 6 7 8 i16x8.all_true call $log)"#, "1" },
    test_i16x8_bitmask => { r#"(func (export "_start")
        v128.const i16x8 0x8000 0 0x8000 0 0 0 0 0 i16x8.bitmask call $log)"#, "5" },

    test_i16x8_extend_low_i8x16_s => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extend_low_i8x16_s i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_extend_low_i8x16_u => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extend_low_i8x16_u i16x8.extract_lane_u 0 call $log)"#, "255" },
    test_i16x8_extend_high_i8x16_s => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0x80 0 0 0 0 0 0 0
        i16x8.extend_high_i8x16_s i16x8.extract_lane_s 0 call $log)"#, "-128" },

    test_i16x8_extadd_pairwise_i8x16_s => { r#"(func (export "_start")
        v128.const i8x16 3 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extadd_pairwise_i8x16_s i16x8.extract_lane_s 0 call $log)"#, "7" },
    test_i16x8_q15mulr_sat_s => { r#"(func (export "_start")
        v128.const i16x8 0x8000 0 0 0 0 0 0 0 v128.const i16x8 0x8000 0 0 0 0 0 0 0
        i16x8.q15mulr_sat_s i16x8.extract_lane_s 0 call $log)"#, "32767" },
    test_i16x8_extmul_low_i8x16_s => { r#"(func (export "_start")
        v128.const i8x16 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 20 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extmul_low_i8x16_s i16x8.extract_lane_s 0 call $log)"#, "200" },
    test_i16x8_narrow_i32x4_s => { r#"(func (export "_start")
        v128.const i32x4 100000 0 0 0 v128.const i32x4 0 0 0 0
        i16x8.narrow_i32x4_s i16x8.extract_lane_s 0 call $log)"#, "32767" },

    // ── remaining comparisons + widening variants ────────────────────────────
    test_i16x8_ne => { r#"(func (export "_start")
        v128.const i16x8 5 0 0 0 0 0 0 0 v128.const i16x8 6 0 0 0 0 0 0 0
        i16x8.ne i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_lt_u => { r#"(func (export "_start")
        v128.const i16x8 1 0 0 0 0 0 0 0 v128.const i16x8 0xFFFF 0 0 0 0 0 0 0
        i16x8.lt_u i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_gt_s => { r#"(func (export "_start")
        v128.const i16x8 5 0 0 0 0 0 0 0 v128.const i16x8 3 0 0 0 0 0 0 0
        i16x8.gt_s i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_le_s => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 0 0 0 0 0 0 0 0
        i16x8.le_s i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_le_u => { r#"(func (export "_start")
        v128.const i16x8 1 0 0 0 0 0 0 0 v128.const i16x8 0xFFFF 0 0 0 0 0 0 0
        i16x8.le_u i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_ge_s => { r#"(func (export "_start")
        v128.const i16x8 5 0 0 0 0 0 0 0 v128.const i16x8 5 0 0 0 0 0 0 0
        i16x8.ge_s i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_ge_u => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0 0 0 0 0 0 0 v128.const i16x8 1 0 0 0 0 0 0 0
        i16x8.ge_u i16x8.extract_lane_s 0 call $log)"#, "-1" },
    test_i16x8_extend_high_i8x16_u => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0xFF 0 0 0 0 0 0 0
        i16x8.extend_high_i8x16_u i16x8.extract_lane_u 0 call $log)"#, "255" },
    test_i16x8_extmul_high_i8x16_s => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 10 0 0 0 0 0 0 0
        v128.const i8x16 0 0 0 0 0 0 0 0 20 0 0 0 0 0 0 0
        i16x8.extmul_high_i8x16_s i16x8.extract_lane_s 0 call $log)"#, "200" },
    test_i16x8_extmul_low_i8x16_u => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 2 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extmul_low_i8x16_u i16x8.extract_lane_u 0 call $log)"#, "510" },
    test_i16x8_narrow_i32x4_u => { r#"(func (export "_start")
        v128.const i32x4 100000 0 0 0 v128.const i32x4 0 0 0 0
        i16x8.narrow_i32x4_u i16x8.extract_lane_u 0 call $log)"#, "65535" },
    test_i16x8_extmul_high_i8x16_u => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0xFF 0 0 0 0 0 0 0
        v128.const i8x16 0 0 0 0 0 0 0 0 2 0 0 0 0 0 0 0
        i16x8.extmul_high_i8x16_u i16x8.extract_lane_u 0 call $log)"#, "510" },
}
