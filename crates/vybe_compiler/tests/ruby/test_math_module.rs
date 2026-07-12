macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_math_pi,
    "puts Math::PI > 3.14 && Math::PI < 3.15",
    "true"
);
ruby_test!(test_math_e, "puts Math::E > 2.71 && Math::E < 2.72", "true");
ruby_test!(test_math_sqrt, "puts Math.sqrt(9)", "3.0");
ruby_test!(test_math_sin, "puts Math.sin(Math::PI / 2).round(2)", "1.0");
ruby_test!(test_math_cos, "puts Math.cos(Math::PI).round(2)", "-1.0");
ruby_test!(test_math_tan, "puts Math.tan(0).round(2)", "0.0");
ruby_test!(test_math_log, "puts Math.log(Math::E).round(2)", "1.0");
ruby_test!(test_math_log10, "puts Math.log10(100).round(2)", "2.0");
ruby_test!(test_math_hypot, "puts Math.hypot(3, 4)", "5.0");
