use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_case_upcase, "puts 'hello'.upcase", "HELLO");
ruby_test!(test_string_case_downcase, "puts 'HELLO'.downcase", "hello");
ruby_test!(test_string_case_capitalize, "puts 'hello'.capitalize", "Hello");
ruby_test!(test_string_case_swapcase, "puts 'Hello'.swapcase", "hELLO");
ruby_test!(test_string_case_upcase_bang, "s = 'hello'; s.upcase!; puts s", "HELLO");
ruby_test!(test_string_case_downcase_bang, "s = 'HELLO'; s.downcase!; puts s", "hello");
ruby_test!(test_string_case_capitalize_bang, "s = 'hello'; s.capitalize!; puts s", "Hello");
ruby_test!(test_string_case_swapcase_bang, "s = 'Hello'; s.swapcase!; puts s", "hELLO");
ruby_test!(test_string_case_upcase_bang_no_change, "s = 'HELLO'; puts s.upcase!.nil?", "true");
ruby_test!(test_string_case_downcase_bang_no_change, "s = 'hello'; puts s.downcase!.nil?", "true");
