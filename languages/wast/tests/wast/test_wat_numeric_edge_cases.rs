//! Numeric edge cases from the WebAssembly spec: shift-count masking, division
//! and remainder sign/trap rules, rotates, and IEEE-754 special values.
//! These complement the basic i32/i64/f32/f64 suites with the boundary behaviour.
use crate::wat_exec;

wat_exec! {
    // ── shift counts are taken modulo the operand width ──────────────────────
    test_i32_shl_count_masked_to_32 => { r#"(func (export "_start")
        i32.const 1 i32.const 32 i32.shl call $log)"#, "1" },
    test_i32_shl_count_33_shifts_by_1 => { r#"(func (export "_start")
        i32.const 1 i32.const 33 i32.shl call $log)"#, "2" },
    test_i32_shr_u_count_masked => { r#"(func (export "_start")
        i32.const 256 i32.const 40 i32.shr_u call $log)"#, "1" },
    test_i64_shl_count_masked_to_64 => { r#"(func (export "_start")
        i64.const 1 i64.const 64 i64.shl call $log_i64)"#, "1" },
    test_i64_shl_count_65_shifts_by_1 => { r#"(func (export "_start")
        i64.const 1 i64.const 65 i64.shl call $log_i64)"#, "2" },

    // ── arithmetic vs logical right shift on a negative value ────────────────
    test_i32_shr_s_sign_extends => { r#"(func (export "_start")
        i32.const -8 i32.const 1 i32.shr_s call $log)"#, "-4" },
    test_i32_shr_u_zero_fills => { r#"(func (export "_start")
        i32.const -1 i32.const 28 i32.shr_u call $log)"#, "15" },

    // ── rotate ───────────────────────────────────────────────────────────────
    test_i32_rotl_wraps_top_bit_to_bottom => { r#"(func (export "_start")
        i32.const 0x80000000 i32.const 1 i32.rotl call $log)"#, "1" },
    test_i32_rotr_wraps_bottom_bit_to_top => { r#"(func (export "_start")
        i32.const 1 i32.const 1 i32.rotr call $log)"#, "-2147483648" },
    test_i64_rotl_full_width => { r#"(func (export "_start")
        i64.const 0x8000000000000000 i64.const 1 i64.rotl call $log_i64)"#, "1" },

    // ── signed division / remainder sign rules ───────────────────────────────
    test_i32_div_s_truncates_toward_zero => { r#"(func (export "_start")
        i32.const -7 i32.const 2 i32.div_s call $log)"#, "-3" },
    test_i32_rem_s_takes_sign_of_dividend => { r#"(func (export "_start")
        i32.const -7 i32.const 2 i32.rem_s call $log)"#, "-1" },
    test_i32_rem_s_positive_dividend_negative_divisor => { r#"(func (export "_start")
        i32.const 7 i32.const -2 i32.rem_s call $log)"#, "1" },
    test_i32_div_u_treats_operands_unsigned => { r#"(func (export "_start")
        i32.const -2 i32.const 2 i32.div_u call $log)"#, "2147483647" },
    test_i32_rem_u_high_bit => { r#"(func (export "_start")
        i32.const -1 i32.const 2 i32.rem_u call $log)"#, "1" },

    // ── division traps ───────────────────────────────────────────────────────
    test_i32_div_s_by_zero_traps => { r#"(func (export "_start")
        i32.const 5 i32.const 0 i32.div_s call $log)"#, "trap" },
    test_i32_div_u_by_zero_traps => { r#"(func (export "_start")
        i32.const 5 i32.const 0 i32.div_u call $log)"#, "trap" },
    test_i32_rem_s_by_zero_traps => { r#"(func (export "_start")
        i32.const 5 i32.const 0 i32.rem_s call $log)"#, "trap" },
    test_i32_div_s_min_by_neg1_traps => { r#"(func (export "_start")
        i32.const -2147483648 i32.const -1 i32.div_s call $log)"#, "trap" },
    test_i64_div_s_min_by_neg1_traps => { r#"(func (export "_start")
        i64.const -9223372036854775808 i64.const -1 i64.div_s call $log_i64)"#, "trap" },

    // ── rem_s of INT_MIN by -1 is defined as 0 (does NOT trap) ────────────────
    test_i32_rem_s_min_by_neg1_is_zero => { r#"(func (export "_start")
        i32.const -2147483648 i32.const -1 i32.rem_s call $log)"#, "0" },

    // ── float special values: NaN, infinities, signed zero ───────────────────
    test_f64_div_by_zero_is_infinity => { r#"(func (export "_start")
        f64.const 1.0 f64.const 0.0 f64.div call $log_f64)"#, "inf" },
    test_f64_neg_div_by_zero_is_neg_infinity => { r#"(func (export "_start")
        f64.const -1.0 f64.const 0.0 f64.div call $log_f64)"#, "-inf" },
    test_f64_zero_div_zero_is_nan => { r#"(func (export "_start")
        f64.const 0.0 f64.const 0.0 f64.div call $log_f64)"#, "nan" },
    test_f64_sqrt_of_negative_is_nan => { r#"(func (export "_start")
        f64.const -4.0 f64.sqrt call $log_f64)"#, "nan" },
    test_f64_inf_minus_inf_is_nan => { r#"(func (export "_start")
        f64.const inf f64.const inf f64.sub call $log_f64)"#, "nan" },

    // ── min/max NaN propagation and signed-zero rules ────────────────────────
    test_f64_min_with_nan_is_nan => { r#"(func (export "_start")
        f64.const 1.0 f64.const nan f64.min call $log_f64)"#, "nan" },
    test_f64_max_with_nan_is_nan => { r#"(func (export "_start")
        f64.const 1.0 f64.const nan f64.max call $log_f64)"#, "nan" },
    test_f32_min_with_nan_is_nan => { r#"(func (export "_start")
        f32.const 1.0 f32.const nan f32.min call $log_f32)"#, "nan" },

    // ── copysign carries the sign of the second operand ──────────────────────
    test_f64_copysign_negative => { r#"(func (export "_start")
        f64.const 3.0 f64.const -1.0 f64.copysign call $log_f64)"#, "-3.0" },
    test_f64_copysign_positive_from_negative => { r#"(func (export "_start")
        f64.const -3.0 f64.const 1.0 f64.copysign call $log_f64)"#, "3.0" },

    // ── nearest uses round-half-to-even ──────────────────────────────────────
    test_f64_nearest_half_rounds_to_even_down => { r#"(func (export "_start")
        f64.const 2.5 f64.nearest call $log_f64)"#, "2.0" },
    test_f64_nearest_half_rounds_to_even_up => { r#"(func (export "_start")
        f64.const 3.5 f64.nearest call $log_f64)"#, "4.0" },
    test_f32_nearest_half_to_even => { r#"(func (export "_start")
        f32.const 0.5 f32.nearest call $log_f32)"#, "0.0" },

    // ── trunc toward zero for negatives ──────────────────────────────────────
    test_f64_trunc_negative_toward_zero => { r#"(func (export "_start")
        f64.const -2.9 f64.trunc call $log_f64)"#, "-2.0" },
    test_f64_floor_negative => { r#"(func (export "_start")
        f64.const -2.1 f64.floor call $log_f64)"#, "-3.0" },
    test_f64_ceil_negative => { r#"(func (export "_start")
        f64.const -2.9 f64.ceil call $log_f64)"#, "-2.0" },

    // ── float→int truncation traps out of range ──────────────────────────────
    test_i32_trunc_f64_s_out_of_range_traps => { r#"(func (export "_start")
        f64.const 1e19 i32.trunc_f64_s call $log)"#, "trap" },
    test_i32_trunc_f64_s_nan_traps => { r#"(func (export "_start")
        f64.const nan i32.trunc_f64_s call $log)"#, "trap" },
    test_i32_trunc_sat_f64_s_clamps_high => { r#"(func (export "_start")
        f64.const 1e19 i32.trunc_sat_f64_s call $log)"#, "2147483647" },
    test_i32_trunc_sat_f64_s_nan_is_zero => { r#"(func (export "_start")
        f64.const nan i32.trunc_sat_f64_s call $log)"#, "0" },
    test_i32_trunc_sat_f64_u_negative_is_zero => { r#"(func (export "_start")
        f64.const -1.0 i32.trunc_sat_f64_u call $log)"#, "0" },
}
