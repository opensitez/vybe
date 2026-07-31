kotlin_run_test!(
    test_int_maximum_overflow_wraps_like_runtime,
    r#"
        fun main() {
            println(Int.MAX_VALUE + 1)
        }
    "#,
    &["-2147483648"]
);

kotlin_run_test!(
    test_int_minimum_overflow_stays_min_when_negating_minimum,
    r#"
        fun main() {
            println(Int.MIN_VALUE * -1)
        }
    "#,
    &["-2147483648"]
);

kotlin_run_test!(
    test_long_max_increment_wraps,
    r#"
        fun main() {
            println(Long.MAX_VALUE + 1)
        }
    "#,
    &["-9223372036854775808"]
);

kotlin_run_test!(
    test_unsigned_conversion_roundtrip,
    r#"
        fun main() {
            val n: Int = -1
            val u = n.toUInt()
            println(u)
            println(u.toInt())
        }
    "#,
    &["4294967295", "-1"]
);

kotlin_run_test!(
    test_float_to_int_truncation_direction,
    r#"
        fun main() {
            println(3.9.toInt())
            println((-3.9).toInt())
        }
    "#,
    &["3", "-3"]
);

kotlin_run_test!(
    test_double_is_infinite_roundtrip,
    r#"
        fun main() {
            val inf = Double.POSITIVE_INFINITY
            println(inf.isInfinite())
            val back = inf / 2
            println(back.isInfinite())
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_nan_roundtrip_boolean,
    r#"
        fun main() {
            val nan = Double.NaN
            println(nan.isNaN())
            println((nan == nan))
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_double_positive_to_short_is_truncated,
    r#"
        fun main() {
            println(128.9.toShort())
            println(-129.1.toShort())
        }
    "#,
    &["-128", "127"]
);

kotlin_run_test!(
    test_parse_int_beyond_range_wraps,
    r#"
        fun main() {
            println("2147483647".toInt())
            println("-2147483648".toLong())
        }
    "#,
    &["2147483647", "-2147483648"]
);

kotlin_run_test!(
    test_math_division_and_mod_boundary,
    r#"
        fun main() {
            println(0 / Int.MAX_VALUE)
            println(Int.MIN_VALUE % -1)
            println(7 % 3)
        }
    "#,
    &["0", "0", "1"]
);
