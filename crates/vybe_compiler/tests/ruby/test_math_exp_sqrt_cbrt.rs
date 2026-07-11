
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_exp_basic, "puts Math.exp(0)", "1.0");
ruby_test!(test_exp_one, "puts Math.exp(1) == Math::E", "true");
ruby_test!(test_exp_negative, "puts Math.exp(-Float::INFINITY)", "0.0");
ruby_test!(test_sqrt_basic, "puts Math.sqrt(9)", "3.0");
ruby_test!(test_sqrt_zero, "puts Math.sqrt(0)", "0.0");
ruby_test!(test_sqrt_negative_domain_error, "begin; Math.sqrt(-1); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_cbrt_basic, "puts Math.cbrt(27)", "3.0");
ruby_test!(test_cbrt_zero, "puts Math.cbrt(0)", "0.0");
ruby_test!(test_cbrt_negative, "puts Math.cbrt(-8)", "-2.0"); // cbrt allows negative input
ruby_test!(test_cbrt_infinity, "puts Math.cbrt(Float::INFINITY)", "Infinity");
