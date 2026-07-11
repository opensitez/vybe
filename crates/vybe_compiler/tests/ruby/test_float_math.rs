
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_float_math_add, "puts 1.5 + 2.5 == 4.0", "true");
ruby_test!(test_float_math_sub, "puts 2.5 - 1.5 == 1.0", "true");
ruby_test!(test_float_math_mul, "puts 1.5 * 2.0 == 3.0", "true");
ruby_test!(test_float_math_div, "puts 3.0 / 2.0 == 1.5", "true");
ruby_test!(test_float_math_mod, "puts 3.5 % 2.0 == 1.5", "true");
ruby_test!(test_float_math_divmod, "puts 3.5.divmod(2.0).join('-')", "1-1.5");
ruby_test!(test_float_math_pow, "puts 2.0 ** 3.0 == 8.0", "true");
ruby_test!(test_float_math_abs, "puts (-1.5).abs == 1.5", "true");
ruby_test!(test_float_math_magnitude, "puts (-1.5).magnitude == 1.5", "true");
ruby_test!(test_float_math_zero, "puts 0.0.zero?", "true");
ruby_test!(test_float_math_zero_false, "puts 1.5.zero?", "false");
ruby_test!(test_float_math_positive, "puts 1.5.positive?", "true");
ruby_test!(test_float_math_negative, "puts (-1.5).negative?", "true");
ruby_test!(test_float_math_finite, "puts 1.5.finite?", "true");
ruby_test!(test_float_math_finite_infinity, "puts Float::INFINITY.finite?", "false");
ruby_test!(test_float_math_infinite, "puts Float::INFINITY.infinite?", "1");
ruby_test!(test_float_math_infinite_negative, "puts (-Float::INFINITY).infinite?", "-1");
ruby_test!(test_float_math_nan, "puts Float::NAN.nan?", "true");
