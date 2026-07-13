//! Numeric literal formats from the WASM text spec (6.1) — decimal, hex,
//! underscores, floats, exponents, hex floats, and the special float tokens.
//! Verified by value so both the lexer and the parsed magnitude are checked.
use crate::wat_exec;

wat_exec! {
    test_decimal_int => { r#"(func (export "_start") i32.const 12345 call $log)"#, "12345" },
    test_negative_int => { r#"(func (export "_start") i32.const -42 call $log)"#, "-42" },
    test_hex_int => { r#"(func (export "_start") i32.const 0xFF call $log)"#, "255" },
    test_hex_int_uppercase => { r#"(func (export "_start") i32.const 0xABCD call $log)"#, "43981" },
    test_underscores_in_decimal => { r#"(func (export "_start") i32.const 1_000_000 call $log)"#, "1000000" },
    test_underscores_in_hex => { r#"(func (export "_start") i32.const 0xFF_FF call $log)"#, "65535" },
    test_i64_large_decimal => { r#"(func (export "_start")
        i64.const 9223372036854775807 call $log_i64)"#, "9223372036854775807" },
    test_i64_hex_full_width => { r#"(func (export "_start")
        i64.const 0xFFFFFFFFFFFFFFFF call $log_i64)"#, "-1" },
    test_i32_hex_high_bit => { r#"(func (export "_start") i32.const 0x80000000 call $log)"#, "-2147483648" },
    test_float_decimal => { r#"(func (export "_start") f64.const 3.14 call $log_f64)"#, "3.14" },
    test_float_negative => { r#"(func (export "_start") f64.const -2.5 call $log_f64)"#, "-2.5" },
    test_float_exponent_positive => { r#"(func (export "_start") f64.const 1.5e3 call $log_f64)"#, "1500.0" },
    test_float_exponent_negative => { r#"(func (export "_start") f64.const 2.5e-1 call $log_f64)"#, "0.25" },
    test_float_capital_exponent => { r#"(func (export "_start") f64.const 1E2 call $log_f64)"#, "100.0" },
    test_float_underscores => { r#"(func (export "_start") f64.const 1_000.5 call $log_f64)"#, "1000.5" },
    test_float_leading_zero => { r#"(func (export "_start") f64.const 0.125 call $log_f64)"#, "0.125" },
    test_hex_float => { r#"(func (export "_start") f64.const 0x1.8p1 call $log_f64)"#, "3.0" },
    test_hex_float_fraction => { r#"(func (export "_start") f64.const 0x1.0p4 call $log_f64)"#, "16.0" },
    test_positive_infinity => { r#"(func (export "_start") f64.const inf call $log_f64)"#, "inf" },
    test_negative_infinity => { r#"(func (export "_start") f64.const -inf call $log_f64)"#, "-inf" },
    test_nan_literal => { r#"(func (export "_start") f64.const nan call $log_f64)"#, "nan" },
    test_nan_with_payload => { r#"(func (export "_start") f64.const nan:0x400000 call $log_f64)"#, "nan" },
    test_f32_literal => { r#"(func (export "_start") f32.const 1.5 call $log_f32)"#, "1.5" },
    test_negative_zero_preserves_sign => { r#"(func (export "_start") f64.const -0.0 call $log_f64)"#, "-0.0" },
    test_large_float_exponent => { r#"(func (export "_start") f64.const 1e100 f64.const 1e100 f64.div call $log_f64)"#, "1.0" },
}
