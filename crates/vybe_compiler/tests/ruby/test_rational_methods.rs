macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_rational_creation,
    "puts Rational(1, 2).class.name",
    "Rational"
);
ruby_test!(
    test_rational_numerator,
    "puts Rational(2, 4).numerator",
    "1"
);
ruby_test!(
    test_rational_denominator,
    "puts Rational(2, 4).denominator",
    "2"
);
ruby_test!(
    test_rational_addition,
    "puts (Rational(1, 2) + Rational(1, 4))",
    "3/4"
);
ruby_test!(
    test_rational_subtraction,
    "puts (Rational(1, 2) - Rational(1, 4))",
    "1/4"
);
ruby_test!(
    test_rational_multiplication,
    "puts (Rational(1, 2) * Rational(1, 4))",
    "1/8"
);
ruby_test!(
    test_rational_division,
    "puts (Rational(1, 2) / Rational(1, 4))",
    "2/1"
);
ruby_test!(test_rational_to_f, "puts Rational(1, 2).to_f", "0.5");
ruby_test!(test_rational_to_s, "puts Rational(1, 2).to_s", "1/2");
ruby_test!(test_rational_to_i, "puts Rational(5, 2).to_i", "2");
ruby_test!(test_rational_rationalize, "puts 0.5.rationalize", "1/2");
ruby_test!(
    test_rational_compare,
    "puts Rational(1, 2) == Rational(2, 4)",
    "true"
);
ruby_test!(
    test_rational_compare_float,
    "puts Rational(1, 2) == 0.5",
    "true"
);
