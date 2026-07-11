
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_math_advanced_acos, "puts (Math.acos(1.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_acosh, "puts (Math.acosh(1.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_asin, "puts (Math.asin(0.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_asinh, "puts (Math.asinh(0.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_atan, "puts (Math.atan(0.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_atan2, "puts (Math.atan2(0.0, 1.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_atanh, "puts (Math.atanh(0.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_cbrt, "puts (Math.cbrt(8.0) * 1000).to_i", "2000");
ruby_test!(test_math_advanced_erf, "puts (Math.erf(0.0) * 1000).to_i", "0");
ruby_test!(test_math_advanced_erfc, "puts (Math.erfc(0.0) * 1000).to_i", "1000");
ruby_test!(test_math_advanced_frexp, "m, e = Math.frexp(8.0); puts \"#{(m * 1000).to_i}-#{e}\"", "500-4");
ruby_test!(test_math_advanced_hypot, "puts (Math.hypot(3.0, 4.0) * 1000).to_i", "5000");
ruby_test!(test_math_advanced_ldexp, "puts (Math.ldexp(0.5, 4) * 1000).to_i", "8000");
ruby_test!(test_math_advanced_lgamma, "a, s = Math.lgamma(1.0); puts \"#{(a * 1000).to_i}-#{s}\"", "0-1");
ruby_test!(test_math_advanced_log2, "puts (Math.log2(8.0) * 1000).to_i", "3000");
