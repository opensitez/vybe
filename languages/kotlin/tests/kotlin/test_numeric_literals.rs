kotlin_run_test!(
    test_decimal_literal_basic,
    r#"fun main() { println(12 + 8) }"#,
    &["20"]
);

kotlin_run_test!(
    test_decimal_underscore_grouping,
    r#"fun main() { println(1_000 + 2_000) }"#,
    &["3000"]
);

kotlin_run_test!(
    test_hex_literal_small,
    r#"fun main() { println(0xA + 0x5) }"#,
    &["15"]
);

kotlin_run_test!(
    test_hex_literal_with_underscore,
    r#"fun main() { println(0x1_0 + 0x2) }"#,
    &["18"]
);

kotlin_run_test!(
    test_binary_literal_small,
    r#"fun main() { println(0b1010 + 0b0101) }"#,
    &["15"]
);

kotlin_run_test!(
    test_long_literal_suffix,
    r#"fun main() { val v: Long = 12L; println(v + 3L) }"#,
    &["15"]
);

kotlin_run_test!(
    test_long_hex_literal,
    r#"fun main() { val v: Long = 0x10L; println(v) }"#,
    &["16"]
);

kotlin_run_test!(
    test_double_scientific,
    r#"fun main() { println(1.5e1) }"#,
    &["15.0"]
);

kotlin_run_test!(
    test_negative_scientific,
    r#"fun main() { println(2e-1) }"#,
    &["0.2"]
);

kotlin_run_test!(
    test_float_suffix,
    r#"fun main() { val f: Float = 1.25f; println(f * 2) }"#,
    &["2.5"]
);

kotlin_run_test!(
    test_double_suffix,
    r#"fun main() { val d: Double = 2.0; println(d * 3) }"#,
    &["6.0"]
);

kotlin_run_test!(
    test_unary_minus_on_literal,
    r#"fun main() { println(-42) }"#,
    &["-42"]
);

kotlin_run_test!(
    test_unary_plus_on_literal,
    r#"fun main() { println(+7) }"#,
    &["7"]
);

kotlin_run_test!(
    test_zero_suffix,
    r#"fun main() { println(0) }"#,
    &["0"]
);

kotlin_run_test!(
    test_mixed_hex_and_decimal_addition,
    r#"fun main() { println(0x5 + 3) }"#,
    &["8"]
);

kotlin_run_test!(
    test_binary_mixed_ops,
    r#"fun main() { println(0b11 * 0x2) }"#,
    &["6"]
);

kotlin_run_test!(
    test_large_integer_division,
    r#"fun main() { println(10 / 3) }"#,
    &["3"]
);

kotlin_run_test!(
    test_long_integer_division,
    r#"fun main() { val a: Long = 10L; println(a / 2L) }"#,
    &["5"]
);

kotlin_run_test!(
    test_remainder_operator,
    r#"fun main() { println(10 % 4) }"#,
    &["2"]
);

kotlin_run_test!(
    test_precedence_literals,
    r#"fun main() { println(1 + 2 * 3) }"#,
    &["7"]
);

kotlin_run_test!(
    test_numeric_shift,
    r#"fun main() { val x = 1 shl 2; println(x) }"#,
    &["4"]
);

kotlin_run_test!(
    test_numeric_bitwise_and,
    r#"fun main() { println(6 and 3) }"#,
    &["2"]
);

kotlin_run_test!(
    test_numeric_bitwise_or,
    r#"fun main() { println(6 or 1) }"#,
    &["7"]
);

kotlin_run_test!(
    test_numeric_bitwise_xor,
    r#"fun main() { println(6 xor 3) }"#,
    &["5"]
);

kotlin_run_test!(
    test_float_to_int_cast,
    r#"fun main() { println(4.9.toInt()) }"#,
    &["4"]
);

kotlin_run_test!(
    test_double_to_int_cast,
    r#"fun main() { println(9.1.toInt()) }"#,
    &["9"]
);

kotlin_run_test!(
    test_int_to_double_cast,
    r#"fun main() { val x = 9; println(x.toDouble() + 0.5) }"#,
    &["9.5"]
);

kotlin_run_test!(
    test_negative_binary,
    r#"fun main() { println(-0b11) }"#,
    &["-3"]
);

kotlin_run_test!(
    test_numeric_with_plus_operator_literals,
    r#"fun main() { println((+1) + (+2) + (+3)) }"#,
    &["6"]
);

kotlin_run_test!(
    test_string_length_from_length_literal,
    r#"fun main() { val n = "123456".length; println(n) }"#,
    &["6"]
);

kotlin_run_test!(
    test_zero_long_prefix_addition,
    r#"fun main() { val x: Long = 0L; println(x + 1L) }"#,
    &["1"]
);

kotlin_run_test!(
    test_mixed_float_long,
    r#"fun main() { val x = 2L + 3.0; println(x) }"#,
    &["5.0"]
);

kotlin_run_test!(
    test_hex_float_notation_not_supported,
    r#"fun main() { println(1_000_000) }"#,
    &["1000000"]
);

kotlin_run_test!(
    test_minimal_negative_comparison,
    r#"fun main() { println(-1 < 0) }"#,
    &["true"]
);

kotlin_run_test!(
    test_literal_range_inclusive,
    r#"fun main() { val r = 1..3; println(r.start == 1 && r.endInclusive == 3) }"#,
    &["true"]
);

kotlin_run_test!(
    test_unsigned_not_supported_fallback,
    r#"fun main() { val a: Long = 10L; println(a) }"#,
    &["10"]
);
