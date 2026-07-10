use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_range_include_question, "puts (1..5).include?(3)", "true");
ruby_test!(test_range_include_question_false, "puts (1..5).include?(6)", "false");
ruby_test!(test_range_include_question_exclusive, "puts (1...5).include?(5)", "false");
ruby_test!(test_range_member_question, "puts (1..5).member?(3)", "true");
ruby_test!(test_range_member_question_string, "puts ('a'..'z').member?('c')", "true");
ruby_test!(test_range_cover_question, "puts (1..5).cover?(3)", "true");
ruby_test!(test_range_cover_question_false, "puts (1..5).cover?(6)", "false");
ruby_test!(test_range_cover_question_exclusive, "puts (1...5).cover?(5)", "false");
ruby_test!(test_range_cover_question_range, "puts (1..5).cover?(2..4)", "true");
ruby_test!(test_range_cover_question_range_false, "puts (1..5).cover?(4..6)", "false");
ruby_test!(test_range_step, "acc = []; (1..5).step(2) { |x| acc << x }; puts acc.join('-')", "1-3-5");
ruby_test!(test_range_step_enumerator, "puts (1..5).step(2).class.name", "Enumerator::ArithmeticSequence");
ruby_test!(test_range_bsearch, "puts (1..10).bsearch { |x| x >= 5 }", "5");
ruby_test!(test_range_bsearch_not_found, "puts (1..10).bsearch { |x| x > 10 }.nil?", "true");
