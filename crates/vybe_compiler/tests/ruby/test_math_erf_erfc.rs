
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_erf_zero, "puts Math.erf(0)", "0.0");
ruby_test!(test_erf_infinity, "puts Math.erf(Float::INFINITY)", "1.0");
ruby_test!(test_erf_negative_infinity, "puts Math.erf(-Float::INFINITY)", "-1.0");
ruby_test!(test_erfc_zero, "puts Math.erfc(0)", "1.0");
ruby_test!(test_erfc_infinity, "puts Math.erfc(Float::INFINITY)", "0.0");
ruby_test!(test_erfc_negative_infinity, "puts Math.erfc(-Float::INFINITY)", "2.0");
ruby_test!(test_erf_erfc_sum, "puts (Math.erf(0.5) + Math.erfc(0.5)).round(5)", "1.0");
ruby_test!(test_erf_nan, "puts Math.erf(Float::NAN).nan?", "true");
ruby_test!(test_erfc_nan, "puts Math.erfc(Float::NAN).nan?", "true");
