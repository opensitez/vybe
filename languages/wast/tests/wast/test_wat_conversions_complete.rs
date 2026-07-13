//! Complete scalar numeric conversion matrix — the int↔float conversions not
//! already covered by test_wat_integer_conversions / test_wat_float_conversions:
//! unsigned convert, i64 truncation (trapping and saturating), and demote/promote.
use crate::wat_exec;

wat_exec! {
    // ── unsigned int → float ─────────────────────────────────────────────────
    test_f32_convert_i32_u_high_bit => { r#"(func (export "_start")
        i32.const -1 f32.convert_i32_u call $log_f32)"#, "4294967300.0" },
    test_f32_convert_i64_s => { r#"(func (export "_start")
        i64.const -1000000 f32.convert_i64_s call $log_f32)"#, "-1000000.0" },
    test_f32_convert_i64_u_high => { r#"(func (export "_start")
        i64.const -1 f32.convert_i64_u call $log_f32)"#, "18446744000000000000.0" },
    test_f64_convert_i32_u_high_bit => { r#"(func (export "_start")
        i32.const -1 f64.convert_i32_u call $log_f64)"#, "4294967295.0" },
    test_f64_convert_i64_s => { r#"(func (export "_start")
        i64.const -5000000000 f64.convert_i64_s call $log_f64)"#, "-5000000000.0" },

    // ── float → i64, trapping trunc ──────────────────────────────────────────
    test_i64_trunc_f32_s => { r#"(func (export "_start")
        f32.const 3.9 i64.trunc_f32_s call $log_i64)"#, "3" },
    test_i64_trunc_f32_u => { r#"(func (export "_start")
        f32.const 100.5 i64.trunc_f32_u call $log_i64)"#, "100" },
    test_i64_trunc_f64_s => { r#"(func (export "_start")
        f64.const -7.9 i64.trunc_f64_s call $log_i64)"#, "-7" },
    test_i64_trunc_f64_u => { r#"(func (export "_start")
        f64.const 9000000000.5 i64.trunc_f64_u call $log_i64)"#, "9000000000" },
    test_i64_trunc_f64_s_out_of_range_traps => { r#"(func (export "_start")
        f64.const 1e19 i64.trunc_f64_s call $log_i64)"#, "trap" },

    // ── float → int, saturating trunc ────────────────────────────────────────
    test_i32_trunc_sat_f32_u_clamps_high => { r#"(func (export "_start")
        f32.const 5e9 i32.trunc_sat_f32_u call $log)"#, "-1" },
    test_i32_trunc_sat_f32_u_negative_is_zero => { r#"(func (export "_start")
        f32.const -3.0 i32.trunc_sat_f32_u call $log)"#, "0" },
    test_i64_trunc_sat_f32_s => { r#"(func (export "_start")
        f32.const 123.9 i64.trunc_sat_f32_s call $log_i64)"#, "123" },
    test_i64_trunc_sat_f32_u_clamps => { r#"(func (export "_start")
        f32.const 1e20 i64.trunc_sat_f32_u call $log_i64)"#, "-1" },
    test_i64_trunc_sat_f64_s_nan_is_zero => { r#"(func (export "_start")
        f64.const nan i64.trunc_sat_f64_s call $log_i64)"#, "0" },
    test_i64_trunc_sat_f64_s_clamps_high => { r#"(func (export "_start")
        f64.const 1e19 i64.trunc_sat_f64_s call $log_i64)"#, "9223372036854775807" },
    test_i64_trunc_sat_f64_u_negative_is_zero => { r#"(func (export "_start")
        f64.const -1.0 i64.trunc_sat_f64_u call $log_i64)"#, "0" },

    // ── demote / promote round-trip ──────────────────────────────────────────
    test_f32_demote_f64_rounds => { r#"(func (export "_start")
        f64.const 1.5 f32.demote_f64 call $log_f32)"#, "1.5" },
    test_f64_promote_f32_exact => { r#"(func (export "_start")
        f32.const 2.5 f64.promote_f32 call $log_f64)"#, "2.5" },
    test_wrap_then_extend_roundtrip => { r#"(func (export "_start")
        i64.const 0x1_0000_002A i32.wrap_i64 i64.extend_i32_s call $log_i64)"#, "42" },
}
