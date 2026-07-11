
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_math_trig_sin, "puts Math.sin(Math::PI / 2).round", "1");
ruby_test!(test_math_trig_cos, "puts Math.cos(Math::PI).round", "-1");
ruby_test!(test_math_trig_tan, "puts Math.tan(0).round", "0");
ruby_test!(test_math_trig_asin, "puts (Math.asin(1) / Math::PI * 2).round", "1");
ruby_test!(test_math_trig_acos, "puts (Math.acos(-1) / Math::PI).round", "1");
ruby_test!(test_math_trig_atan, "puts (Math.atan(1) * 4 / Math::PI).round", "1");
ruby_test!(test_math_trig_atan2, "puts (Math.atan2(1, 1) * 4 / Math::PI).round", "1");
ruby_test!(test_math_trig_sinh, "puts Math.sinh(0)", "0.0");
ruby_test!(test_math_trig_cosh, "puts Math.cosh(0)", "1.0");
ruby_test!(test_math_trig_tanh, "puts Math.tanh(0)", "0.0");
ruby_test!(test_math_trig_asinh, "puts Math.asinh(0)", "0.0");
ruby_test!(test_math_trig_acosh, "puts Math.acosh(1)", "0.0");
ruby_test!(test_math_trig_atanh, "puts Math.atanh(0)", "0.0");
