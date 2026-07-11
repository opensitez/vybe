
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_step_basic, "acc = []; 1.step(5, 2) {|x| acc << x}; puts acc.join('-')", "1-3-5");
ruby_test!(test_step_negative, "acc = []; 5.step(1, -2) {|x| acc << x}; puts acc.join('-')", "5-3-1");
ruby_test!(test_step_float, "acc = []; 1.step(2, 0.5) {|x| acc << x}; puts acc.join('-')", "1.0-1.5-2.0");
ruby_test!(test_step_no_block, "puts 1.step(5, 2).is_a?(Enumerator)", "true");
ruby_test!(test_step_default_step, "acc = []; 1.step(3) {|x| acc << x}; puts acc.join('-')", "1-2-3");
ruby_test!(test_step_zero_step_error, "begin; 1.step(5, 0); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_step_infinite, "acc = []; 1.step {|x| acc << x; break if x > 2}; puts acc.join('-')", "1-2"); // ruby 2.6+ step with no limit
ruby_test!(test_step_keyword_args, "acc = []; 1.step(by: 2, to: 5) {|x| acc << x}; puts acc.join('-')", "1-3-5");
ruby_test!(test_step_arithmetic_sequence, "puts 1.step(5, 2).class.name", "Enumerator::ArithmeticSequence"); // ruby 2.6+
