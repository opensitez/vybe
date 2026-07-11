
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_nonzero_basic, "puts 5.nonzero?", "5");
ruby_test!(test_nonzero_zero, "puts 0.nonzero?.nil?", "true");
ruby_test!(test_nonzero_float, "puts 5.0.nonzero?", "5.0");
ruby_test!(test_nonzero_float_zero, "puts 0.0.nonzero?.nil?", "true");
ruby_test!(test_nonzero_negative, "puts (-5).nonzero?", "-5");
ruby_test!(test_zero_predicate_basic, "puts 0.zero?", "true");
ruby_test!(test_zero_predicate_false, "puts 5.zero?", "false");
ruby_test!(test_zero_predicate_float, "puts 0.0.zero?", "true");
ruby_test!(test_zero_predicate_float_false, "puts 5.0.zero?", "false");
ruby_test!(test_zero_predicate_negative_zero, "puts (-0.0).zero?", "true");
ruby_test!(test_zero_predicate_rational, "puts Rational(0, 1).zero?", "true");
ruby_test!(test_nonzero_rational, "puts Rational(1, 2).nonzero?", "1/2");
