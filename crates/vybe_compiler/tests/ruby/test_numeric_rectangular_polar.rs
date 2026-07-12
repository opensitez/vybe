macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_rectangular_basic, "puts 1.rect.join('-')", "1-0");
ruby_test!(
    test_rectangular_alias,
    "puts 1.rectangular.join('-')",
    "1-0"
);
ruby_test!(test_rectangular_float, "puts 1.5.rect.join('-')", "1.5-0.0");
ruby_test!(
    test_rectangular_negative,
    "puts (-1.5).rect.join('-')",
    "-1.5-0.0"
);
ruby_test!(
    test_rectangular_complex,
    "puts Complex(1, 2).rect.join('-')",
    "1-2"
);
ruby_test!(test_polar_basic, "puts 1.polar.join('-')", "1-0");
ruby_test!(test_polar_float, "puts 1.5.polar.join('-')", "1.5-0");
ruby_test!(
    test_polar_negative,
    "puts (-1).polar.map(&:to_s).join('-')",
    "1-3.141592653589793"
); // Math::PI
ruby_test!(
    test_polar_negative_float,
    "puts (-1.5).polar.map(&:to_s).join('-')",
    "1.5-3.141592653589793"
);
ruby_test!(
    test_polar_complex,
    "puts Complex(0, 1).polar.join('-')",
    "1-1.5707963267948966"
); // Math::PI / 2
ruby_test!(test_real_predicate, "puts 1.real?", "true");
ruby_test!(
    test_real_predicate_complex,
    "puts Complex(1, 2).real?",
    "false"
);
ruby_test!(
    test_real_predicate_complex_zero_imag,
    "puts Complex(1, 0).real?",
    "false"
); // Complex is never real?
