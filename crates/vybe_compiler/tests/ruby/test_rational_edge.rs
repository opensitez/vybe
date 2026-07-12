macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_rational_edge_zero_denominator,
    "begin; Rational(1, 0); rescue ZeroDivisionError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_rational_edge_string_arg,
    "puts Rational('1/3') == Rational(1, 3)",
    "true"
);
ruby_test!(
    test_rational_edge_float_arg,
    "puts Rational(0.5) == Rational(1, 2)",
    "true"
);
ruby_test!(
    test_rational_edge_rational_arg,
    "puts Rational(Rational(1, 2)) == Rational(1, 2)",
    "true"
);
ruby_test!(
    test_rational_edge_simplify,
    "puts Rational(2, 4) == Rational(1, 2)",
    "true"
);
ruby_test!(
    test_rational_edge_negative_den,
    "puts Rational(1, -2) == Rational(-1, 2)",
    "true"
);
ruby_test!(test_rational_edge_to_f, "puts Rational(1, 2).to_f", "0.5");
ruby_test!(test_rational_edge_to_i, "puts Rational(3, 2).to_i", "1");
ruby_test!(
    test_rational_edge_to_r,
    "r = Rational(1, 2); puts r.to_r.equal?(r)",
    "true"
);
ruby_test!(
    test_rational_edge_hash,
    "puts Rational(1, 2).hash == Rational(1, 2).hash",
    "true"
);
