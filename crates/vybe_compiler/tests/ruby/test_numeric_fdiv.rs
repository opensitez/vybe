use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_fdiv_basic, "puts 10.fdiv(3)", "3.3333333333333335");
ruby_test!(test_fdiv_float, "puts 10.0.fdiv(3.0)", "3.3333333333333335");
ruby_test!(test_fdiv_zero, "puts 10.fdiv(0)", "Infinity");
ruby_test!(test_fdiv_negative_zero, "puts 10.fdiv(-0.0)", "-Infinity");
ruby_test!(test_fdiv_zero_by_zero, "puts 0.fdiv(0).nan?", "true");
ruby_test!(test_fdiv_infinity, "puts 10.fdiv(Float::INFINITY)", "0.0");
ruby_test!(test_fdiv_nan, "puts 10.fdiv(Float::NAN).nan?", "true");
ruby_test!(test_fdiv_rational, "puts Rational(1, 2).fdiv(3)", "0.16666666666666666");
ruby_test!(test_fdiv_complex, "puts Complex(1, 2).fdiv(2)", "0.5+1.0i");
