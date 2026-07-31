kotlin_run_test!(
    test_signed_shift_left_behavior,
    r#"
        fun main() {
            println(1 shl 30)
            println(-1 shl 1)
        }
    "#,
    &["1073741824", "-2"]
);

kotlin_run_test!(
    test_signed_shift_right_with_negative,
    r#"
        fun main() {
            println((-4) shr 1)
            println((-4) ushr 1)
        }
    "#,
    &["-2", "2147483646"]
);

kotlin_run_test!(
    test_bitwise_and_masks,
    r#"
        fun main() {
            println(0b1010 and 0b0111)
            println(0b1010 or 0b0101)
            println(0b1010 xor 0b1111)
        }
    "#,
    &["2", "15", "5"]
);

kotlin_run_test!(
    test_integer_division_remainder_contract,
    r#"
        fun main() {
            println(7 / 2)
            println(7 % 2)
            println((-7) / 2)
            println((-7) % 2)
        }
    "#,
    &["3", "1", "-3", "-1"]
);

kotlin_run_test!(
    test_long_shift_and_overflow,
    r#"
        fun main() {
            println(1L shl 63)
            println((-1L) ushr 1)
        }
    "#,
    &["-9223372036854775808", "9223372036854775807"]
);

kotlin_run_test!(
    test_unsigned_mirror_checks,
    r#"
        fun main() {
            println(255u.toByte().toInt())
            println((-1).toUInt())
        }
    "#,
    &["-1", "4294967295"]
);

kotlin_run_test!(
    test_mixed_numeric_precedence_with_overflow,
    r#"
        fun main() {
            val a = Int.MAX_VALUE
            val b = a + 1 - 1
            println(b)
        }
    "#,
    &["2147483647"]
);

kotlin_run_test!(
    test_unary_minus_min_bound,
    r#"
        fun main() {
            println(-Int.MIN_VALUE)
        }
    "#,
    &["-2147483648"]
);

kotlin_run_test!(
    test_power_by_repeated_multiplication,
    r#"
        fun main() {
            val x = 2 * 3 * 4
            println(x)
            val y = 2L * 3L * 4L
            println(y)
        }
    "#,
    &["24", "24"]
);

kotlin_run_test!(
    test_float_precision_overflow,
    r#"
        fun main() {
            val x = Float.MAX_VALUE
            println((x * 2).isInfinite())
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_char_codepoint_arithmetic,
    r#"
        fun main() {
            println('A'.code + 1)
            println(('A'.code + 1).toChar())
        }
    "#,
    &["66", "B"]
);

kotlin_run_test!(
    test_boolean_not_and_xor,
    r#"
        fun main() {
            val a = true
            val b = false
            println(!a)
            println(a xor b)
        }
    "#,
    &["false", "true"]
);
