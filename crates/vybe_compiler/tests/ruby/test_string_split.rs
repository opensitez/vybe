use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_split_basic, "puts 'a,b,c'.split(',').join('-')", "a-b-c");
ruby_test!(test_string_split_space, "puts 'a b c'.split.join('-')", "a-b-c");
ruby_test!(test_string_split_regex, "puts 'a1b2c'.split(/[0-9]/).join('-')", "a-b-c");
ruby_test!(test_string_split_limit, "puts 'a,b,c'.split(',', 2).join('-')", "a-b,c");
ruby_test!(test_string_split_negative_limit, "puts 'a,b,c,,'.split(',', -1).join('-')", "a-b-c--");
ruby_test!(test_string_split_chars, "puts 'abc'.split('').join('-')", "a-b-c");
ruby_test!(test_string_split_block, "acc = []; 'a,b,c'.split(',') { |s| acc << s }; puts acc.join('-')", "a-b-c");
ruby_test!(test_string_split_regex_capture, "puts 'a1b2c'.split(/([0-9])/).join('-')", "a-1-b-2-c");
ruby_test!(test_string_split_empty_regex, "puts 'abc'.split(//).join('-')", "a-b-c");
ruby_test!(test_string_split_awk_space, "puts ' a  b   c '.split(' ').join('-')", "a-b-c");
ruby_test!(test_string_split_exact_space, "puts ' a  b   c '.split(/ /).join('-')", "-a--b---c-"); // wait, space regex splits exactly
ruby_test!(test_string_split_null_byte, "puts \"a\\0b\\0c\".split(\"\\0\").join('-')", "a-b-c");
ruby_test!(test_string_split_trailing_empty, "puts 'a,b,c,,'.split(',').join('-')", "a-b-c"); // trailing empty strings are omitted
