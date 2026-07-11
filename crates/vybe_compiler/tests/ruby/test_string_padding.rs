
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_padding_ljust, "puts 'abc'.ljust(5)", "abc  ");
ruby_test!(test_padding_ljust_char, "puts 'abc'.ljust(5, '*')", "abc**");
ruby_test!(test_padding_ljust_multi_char, "puts 'abc'.ljust(7, 'xy')", "abcxyxy");
ruby_test!(test_padding_ljust_short, "puts 'abc'.ljust(2)", "abc");
ruby_test!(test_padding_rjust, "puts 'abc'.rjust(5)", "  abc");
ruby_test!(test_padding_rjust_char, "puts 'abc'.rjust(5, '*')", "**abc");
ruby_test!(test_padding_rjust_multi_char, "puts 'abc'.rjust(7, 'xy')", "xyxyabc");
ruby_test!(test_padding_rjust_short, "puts 'abc'.rjust(2)", "abc");
ruby_test!(test_padding_center, "puts 'abc'.center(5)", " abc ");
ruby_test!(test_padding_center_char, "puts 'abc'.center(5, '*')", "*abc*");
ruby_test!(test_padding_center_even, "puts 'abc'.center(6, '*')", "*abc**"); // Right gets extra if padding is odd
ruby_test!(test_padding_center_multi, "puts 'abc'.center(7, 'xy')", "xyabcxy");
