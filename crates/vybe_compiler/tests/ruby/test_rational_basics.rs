
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_rational_basic, "puts Rational(1, 2).to_s", "1/2");
ruby_test!(test_rational_reduce, "puts Rational(2, 4).to_s", "1/2");
ruby_test!(test_rational_numerator, "puts Rational(3, 4).numerator", "3");
ruby_test!(test_rational_denominator, "puts Rational(3, 4).denominator", "4");
ruby_test!(test_rational_add, "puts (Rational(1, 2) + Rational(1, 4)).to_s", "3/4");
ruby_test!(test_rational_sub, "puts (Rational(1, 2) - Rational(1, 4)).to_s", "1/4");
ruby_test!(test_rational_mul, "puts (Rational(1, 2) * Rational(1, 4)).to_s", "1/8");
ruby_test!(test_rational_div, "puts (Rational(1, 2) / Rational(1, 4)).to_s", "2/1");
ruby_test!(test_rational_to_f, "puts Rational(1, 2).to_f", "0.5");
