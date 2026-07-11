
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_index_basic, "puts 'hello'.index('l')", "2");
ruby_test!(test_string_index_offset, "puts 'hello'.index('l', 3)", "3");
ruby_test!(test_string_index_regex, "puts 'hello'.index(/[aeiou]/)", "1");
ruby_test!(test_string_index_not_found, "puts 'hello'.index('x').nil?", "true");
ruby_test!(test_string_index_negative_offset, "puts 'hello'.index('l', -3)", "2"); // offset -3 is 'l', so it finds the first 'l' from index 2 onwards, which is 2
ruby_test!(test_string_rindex_basic, "puts 'hello'.rindex('l')", "3");
ruby_test!(test_string_rindex_offset, "puts 'hello'.rindex('l', 2)", "2");
ruby_test!(test_string_rindex_regex, "puts 'hello'.rindex(/[aeiou]/)", "4");
ruby_test!(test_string_rindex_not_found, "puts 'hello'.rindex('x').nil?", "true");
ruby_test!(test_string_rindex_negative_offset, "puts 'hello'.rindex('l', -3)", "2"); // offset -3 is 'l' (index 2), searches backwards, finds 'l' at 2
ruby_test!(test_string_index_empty_string, "puts 'hello'.index('')", "0");
ruby_test!(test_string_rindex_empty_string, "puts 'hello'.rindex('')", "5");
