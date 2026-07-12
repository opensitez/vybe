macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_math_log_basic, "puts Math.log(Math::E).round", "1");
ruby_test!(test_math_log_base, "puts Math.log(100, 10).round", "2");
ruby_test!(test_math_log10, "puts Math.log10(100).round", "2");
ruby_test!(test_math_log2, "puts Math.log2(8).round", "3");
ruby_test!(test_math_exp, "puts Math.exp(1).round(1)", "2.7");
ruby_test!(test_math_sqrt, "puts Math.sqrt(9)", "3.0");
ruby_test!(test_math_cbrt, "puts Math.cbrt(27)", "3.0");
ruby_test!(test_math_hypot, "puts Math.hypot(3, 4)", "5.0");
ruby_test!(test_math_frexp, "puts Math.frexp(128).class.name", "Array");
ruby_test!(test_math_ldexp, "puts Math.ldexp(1.0, 7)", "128.0");
ruby_test!(test_math_erf, "puts Math.erf(0)", "0.0");
ruby_test!(test_math_erfc, "puts Math.erfc(0)", "1.0");
ruby_test!(test_math_gamma, "puts Math.gamma(5)", "24.0");
ruby_test!(test_math_lgamma, "puts Math.lgamma(5).class.name", "Array");
