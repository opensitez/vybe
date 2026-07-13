//! SIMD i8x16 lane operations — 16 lanes of 8-bit integers.
//! Expected values derived from the WebAssembly SIMD spec: 8-bit wrapping,
//! signed/unsigned saturation, and signed/unsigned lane extraction.
use crate::wat_exec;

wat_exec! {
    // ── splat / extract / replace ────────────────────────────────────────────
    test_i8x16_splat_broadcasts => { r#"(func (export "_start")
        i32.const 7 i8x16.splat i8x16.extract_lane_s 9 call $log)"#, "7" },
    test_i8x16_splat_truncates_to_byte => { r#"(func (export "_start")
        i32.const 300 i8x16.splat i8x16.extract_lane_u 0 call $log)"#, "44" },
    test_i8x16_extract_lane_s_sign_extends => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_extract_lane_u_zero_extends => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.extract_lane_u 0 call $log)"#, "255" },
    test_i8x16_replace_lane => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i32.const 42 i8x16.replace_lane 5 i8x16.extract_lane_s 5 call $log)"#, "42" },

    // ── add / sub with 8-bit wrapping ────────────────────────────────────────
    test_i8x16_add_wraps => { r#"(func (export "_start")
        v128.const i8x16 127 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.add i8x16.extract_lane_s 0 call $log)"#, "-128" },
    test_i8x16_sub_wraps_negative => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.sub i8x16.extract_lane_u 0 call $log)"#, "255" },
    test_i8x16_neg => { r#"(func (export "_start")
        v128.const i8x16 100 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.neg i8x16.extract_lane_s 0 call $log)"#, "-100" },
    test_i8x16_neg_min_wraps => { r#"(func (export "_start")
        v128.const i8x16 128 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.neg i8x16.extract_lane_s 0 call $log)"#, "-128" },
    test_i8x16_abs => { r#"(func (export "_start")
        v128.const i8x16 0xFB 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.abs i8x16.extract_lane_s 0 call $log)"#, "5" },

    // ── saturating add/sub, signed and unsigned ──────────────────────────────
    test_i8x16_add_sat_s_clamps_high => { r#"(func (export "_start")
        v128.const i8x16 127 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.add_sat_s i8x16.extract_lane_s 0 call $log)"#, "127" },
    test_i8x16_add_sat_s_clamps_low => { r#"(func (export "_start")
        v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.add_sat_s i8x16.extract_lane_s 0 call $log)"#, "-128" },
    test_i8x16_add_sat_u_clamps => { r#"(func (export "_start")
        v128.const i8x16 250 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.add_sat_u i8x16.extract_lane_u 0 call $log)"#, "255" },
    test_i8x16_sub_sat_s_clamps_low => { r#"(func (export "_start")
        v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.sub_sat_s i8x16.extract_lane_s 0 call $log)"#, "-128" },
    test_i8x16_sub_sat_u_clamps_zero => { r#"(func (export "_start")
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.sub_sat_u i8x16.extract_lane_u 0 call $log)"#, "0" },

    // ── min/max, signed vs unsigned ──────────────────────────────────────────
    test_i8x16_min_s_picks_signed_min => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.min_s i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_min_u_treats_ff_as_255 => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.min_u i8x16.extract_lane_u 0 call $log)"#, "1" },
    test_i8x16_max_s => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.max_s i8x16.extract_lane_s 0 call $log)"#, "1" },
    test_i8x16_max_u => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.max_u i8x16.extract_lane_u 0 call $log)"#, "255" },

    // ── shifts (shift amount taken mod 8) ────────────────────────────────────
    test_i8x16_shl => { r#"(func (export "_start")
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i32.const 3 i8x16.shl i8x16.extract_lane_s 0 call $log)"#, "8" },
    test_i8x16_shr_s_arithmetic => { r#"(func (export "_start")
        v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i32.const 1 i8x16.shr_s i8x16.extract_lane_s 0 call $log)"#, "-64" },
    test_i8x16_shr_u_logical => { r#"(func (export "_start")
        v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i32.const 1 i8x16.shr_u i8x16.extract_lane_u 0 call $log)"#, "64" },

    // ── avgr_u (rounding average) ────────────────────────────────────────────
    test_i8x16_avgr_u_rounds_up => { r#"(func (export "_start")
        v128.const i8x16 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.avgr_u i8x16.extract_lane_u 0 call $log)"#, "4" },

    // ── popcnt (per-lane bit count) ──────────────────────────────────────────
    test_i8x16_popcnt_full_byte => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.popcnt i8x16.extract_lane_u 0 call $log)"#, "8" },
    test_i8x16_popcnt_nibble => { r#"(func (export "_start")
        v128.const i8x16 0x0F 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.popcnt i8x16.extract_lane_u 0 call $log)"#, "4" },

    // ── lane-wise comparisons produce 0x00 / 0xFF lanes ──────────────────────
    test_i8x16_eq_true_lane_is_all_ones => { r#"(func (export "_start")
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.eq i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_eq_false_lane_is_zero => { r#"(func (export "_start")
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 6 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.eq i8x16.extract_lane_s 0 call $log)"#, "0" },
    test_i8x16_lt_s_signed => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.lt_s i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_lt_u_unsigned => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.lt_u i8x16.extract_lane_s 0 call $log)"#, "0" },
    test_i8x16_gt_s => { r#"(func (export "_start")
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.gt_s i8x16.extract_lane_s 0 call $log)"#, "-1" },

    // ── reductions: any_true / all_true / bitmask ────────────────────────────
    test_v128_any_true_detects_nonzero => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0
        v128.any_true call $log)"#, "1" },
    test_v128_any_true_all_zero => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.any_true call $log)"#, "0" },
    test_i8x16_all_true_when_no_zero_lane => { r#"(func (export "_start")
        v128.const i8x16 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16
        i8x16.all_true call $log)"#, "1" },
    test_i8x16_all_true_false_with_zero_lane => { r#"(func (export "_start")
        v128.const i8x16 1 2 3 4 5 6 7 0 9 10 11 12 13 14 15 16
        i8x16.all_true call $log)"#, "0" },
    test_i8x16_bitmask_gathers_sign_bits => { r#"(func (export "_start")
        v128.const i8x16 0x80 0 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.bitmask call $log)"#, "5" },

    // ── narrow (i16x8 → i8x16 with saturation) ───────────────────────────────
    test_i8x16_narrow_i16x8_s_saturates => { r#"(func (export "_start")
        v128.const i16x8 200 -200 0 0 0 0 0 0
        v128.const i16x8 0 0 0 0 0 0 0 0
        i8x16.narrow_i16x8_s i8x16.extract_lane_s 0 call $log)"#, "127" },
    test_i8x16_narrow_i16x8_u_saturates => { r#"(func (export "_start")
        v128.const i16x8 300 0 0 0 0 0 0 0
        v128.const i16x8 0 0 0 0 0 0 0 0
        i8x16.narrow_i16x8_u i8x16.extract_lane_u 0 call $log)"#, "255" },

    // ── swizzle (byte selection by index) ────────────────────────────────────
    test_i8x16_swizzle_selects_bytes => { r#"(func (export "_start")
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 5 15 25 35
        v128.const i8x16 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.swizzle i8x16.extract_lane_u 0 call $log)"#, "40" },
    test_i8x16_swizzle_out_of_range_is_zero => { r#"(func (export "_start")
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 5 15 25 35
        v128.const i8x16 16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.swizzle i8x16.extract_lane_u 0 call $log)"#, "0" },

    // ── remaining comparison variants (ne, unsigned/signed le/ge/gt) ─────────
    test_i8x16_ne => { r#"(func (export "_start")
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 6 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.ne i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_gt_u => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.gt_u i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_le_s => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.le_s i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_le_u => { r#"(func (export "_start")
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.le_u i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_ge_s => { r#"(func (export "_start")
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.ge_s i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_ge_u => { r#"(func (export "_start")
        v128.const i8x16 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.ge_u i8x16.extract_lane_s 0 call $log)"#, "-1" },
    test_i8x16_extadd_pairwise_u => { r#"(func (export "_start")
        v128.const i8x16 200 100 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extadd_pairwise_i8x16_u i16x8.extract_lane_u 0 call $log)"#, "300" },
}
