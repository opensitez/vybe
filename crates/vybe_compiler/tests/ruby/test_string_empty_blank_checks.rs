use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_empty_true, "puts ''.empty?", "true");
ruby_test!(test_empty_false, "puts 'a'.empty?", "false");
ruby_test!(test_empty_space, "puts ' '.empty?", "false");
ruby_test!(test_empty_newline, "puts \"\\n\".empty?", "false");
ruby_test!(test_empty_null_byte, "puts \"\\x00\".empty?", "false");
// Ruby doesn't have .blank? in standard library (it's ActiveSupport), but we can test nil? which is often grouped
ruby_test!(test_nil_question, "puts ''.nil?", "false");
ruby_test!(test_empty_after_clear, "s = 'a'; s.clear; puts s.empty?", "true");
ruby_test!(test_empty_after_gsub, "puts 'a'.gsub('a', '').empty?", "true");
ruby_test!(test_empty_after_strip, "puts ' '.strip.empty?", "true");
ruby_test!(test_empty_after_chomp, "puts \"\\n\".chomp.empty?", "true");
ruby_test!(test_empty_after_chop, "puts 'a'.chop.empty?", "true");
ruby_test!(test_length_zero_is_empty, "s = ''; puts s.length == 0 && s.empty?", "true");
