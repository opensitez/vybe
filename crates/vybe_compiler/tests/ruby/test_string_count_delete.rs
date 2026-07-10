use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_count_basic, "puts 'hello'.count('l')", "2");
ruby_test!(test_string_count_multiple, "puts 'hello'.count('lo')", "3");
ruby_test!(test_string_count_range, "puts 'hello'.count('a-f')", "1");
ruby_test!(test_string_count_negation, "puts 'hello'.count('^l')", "3");
ruby_test!(test_string_delete_basic, "puts 'hello'.delete('l')", "heo");
ruby_test!(test_string_delete_multiple, "puts 'hello'.delete('lo')", "he");
ruby_test!(test_string_delete_range, "puts 'hello'.delete('a-f')", "hllo");
ruby_test!(test_string_delete_negation, "puts 'hello'.delete('^l')", "ll");
ruby_test!(test_string_delete_bang, "s = 'hello'; s.delete!('l'); puts s", "heo");
