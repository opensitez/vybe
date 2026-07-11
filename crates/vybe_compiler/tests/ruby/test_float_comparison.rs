
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_float_comparison_eq, "puts 1.5 == 1.5", "true");
ruby_test!(test_float_comparison_not_eq, "puts 1.5 == 2.5", "false");
ruby_test!(test_float_comparison_greater, "puts 2.5 > 1.5", "true");
ruby_test!(test_float_comparison_less, "puts 1.5 < 2.5", "true");
ruby_test!(test_float_comparison_greater_eq, "puts 2.5 >= 2.5", "true");
ruby_test!(test_float_comparison_less_eq, "puts 1.5 <= 1.5", "true");
ruby_test!(test_float_comparison_cmp, "puts 2.5 <=> 1.5", "1");
ruby_test!(test_float_comparison_cmp_eq, "puts 1.5 <=> 1.5", "0");
ruby_test!(test_float_comparison_cmp_less, "puts 1.5 <=> 2.5", "-1");
ruby_test!(test_float_comparison_nan_cmp, "puts (Float::NAN <=> 1.5).nil?", "true");
ruby_test!(test_float_comparison_nan_eq, "puts Float::NAN == Float::NAN", "false");
ruby_test!(test_float_comparison_infinity, "puts Float::INFINITY > 1e100", "true");
