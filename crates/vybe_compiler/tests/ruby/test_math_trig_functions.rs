use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_sin_zero, "puts Math.sin(0)", "0.0");
ruby_test!(test_sin_pi_half, "puts Math.sin(Math::PI / 2)", "1.0");
ruby_test!(test_cos_zero, "puts Math.cos(0)", "1.0");
ruby_test!(test_cos_pi, "puts Math.cos(Math::PI)", "-1.0");
ruby_test!(test_tan_zero, "puts Math.tan(0)", "0.0");
ruby_test!(test_tan_pi_quarter, "puts Math.tan(Math::PI / 4).round(5)", "1.0");
ruby_test!(test_asin_zero, "puts Math.asin(0)", "0.0");
ruby_test!(test_asin_one, "puts Math.asin(1) == Math::PI / 2", "true");
ruby_test!(test_acos_one, "puts Math.acos(1)", "0.0");
ruby_test!(test_acos_minus_one, "puts Math.acos(-1) == Math::PI", "true");
ruby_test!(test_atan_zero, "puts Math.atan(0)", "0.0");
ruby_test!(test_atan_one, "puts Math.atan(1) == Math::PI / 4", "true");
ruby_test!(test_atan2_zero, "puts Math.atan2(0, 1)", "0.0");
ruby_test!(test_atan2_quadrant_1, "puts Math.atan2(1, 1) == Math::PI / 4", "true");
ruby_test!(test_asin_domain_error, "begin; Math.asin(2); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_acos_domain_error, "begin; Math.acos(2); rescue Math::DomainError; puts 'err'; end", "err");
