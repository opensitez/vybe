macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_rational_creation_basic, "puts Rational(1, 2)", "1/2");
ruby_test!(test_rational_creation_float, "puts Rational(0.5)", "1/2");
ruby_test!(test_rational_creation_string, "puts Rational('1/2')", "1/2");
ruby_test!(
    test_rational_numerator_denominator,
    "r = Rational(3, 4); puts \"#{r.numerator}-#{r.denominator}\"",
    "3-4"
);
ruby_test!(test_rational_simplification, "puts Rational(2, 4)", "1/2");
ruby_test!(
    test_rational_arithmetic_add,
    "puts Rational(1, 2) + Rational(1, 4)",
    "3/4"
);
ruby_test!(
    test_rational_arithmetic_mul,
    "puts Rational(1, 2) * Rational(1, 4)",
    "1/8"
);
ruby_test!(
    test_rational_arithmetic_div,
    "puts Rational(1, 2) / Rational(1, 4)",
    "2/1"
);
ruby_test!(
    test_rational_arithmetic_int,
    "puts Rational(1, 2) + 1",
    "3/2"
);
ruby_test!(test_rational_to_f, "puts Rational(1, 2).to_f", "0.5");
ruby_test!(test_rational_to_i, "puts Rational(5, 2).to_i", "2");
ruby_test!(
    test_rational_zero_division,
    "begin; Rational(1, 0); rescue ZeroDivisionError; puts 'err'; end",
    "err"
);
