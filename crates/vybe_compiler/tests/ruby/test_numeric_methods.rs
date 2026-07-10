use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_numeric_abs, "puts (-42).abs", "42");
ruby_test!(test_numeric_abs_float, "puts (-42.5).abs", "42.5");
ruby_test!(test_numeric_abs_rational, "puts Rational(-1, 2).abs", "1/2");
ruby_test!(test_numeric_abs_complex, "puts Complex(3, 4).abs", "5.0");
ruby_test!(test_numeric_abs2, "puts (-5).abs2", "25");
ruby_test!(test_numeric_abs2_complex, "puts Complex(1, 2).abs2", "5");
ruby_test!(test_numeric_magnitude, "puts (-42).magnitude", "42");
ruby_test!(test_numeric_zero_question, "puts 0.zero?", "true");
ruby_test!(test_numeric_zero_question_false, "puts 1.zero?", "false");
ruby_test!(test_numeric_nonzero_question, "puts 1.nonzero?", "1");
ruby_test!(test_numeric_nonzero_question_false, "puts 0.nonzero?.nil?", "true");
ruby_test!(test_numeric_positive_question, "puts 1.positive?", "true");
ruby_test!(test_numeric_negative_question, "puts (-1).negative?", "true");
ruby_test!(test_numeric_real_question, "puts 1.real?", "true");
ruby_test!(test_numeric_real_question_complex, "puts Complex(1, 2).real?", "false");
ruby_test!(test_numeric_integer_question, "puts 1.integer?", "true");
ruby_test!(test_numeric_integer_question_float, "puts 1.0.integer?", "false");
ruby_test!(test_numeric_step, "acc = []; 1.step(5, 2) { |i| acc << i }; puts acc.join('-')", "1-3-5");
ruby_test!(test_numeric_step_enumerator, "puts 1.step(5).class.name", "Enumerator");
