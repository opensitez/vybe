
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_index_basic, "puts 'hello'.index('e')", "1");
ruby_test!(test_index_regex, "puts 'hello'.index(/[eo]/)", "1");
ruby_test!(test_index_offset, "puts 'hello'.index('l', 3)", "3");
ruby_test!(test_index_offset_negative, "puts 'hello'.index('l', -3)", "2");
ruby_test!(test_index_not_found, "puts 'hello'.index('z').nil?", "true");
ruby_test!(test_index_empty_string, "puts 'hello'.index('')", "0");
ruby_test!(test_index_empty_offset, "puts 'hello'.index('', 2)", "2");
ruby_test!(test_rindex_basic, "puts 'hello'.rindex('l')", "3");
ruby_test!(test_rindex_regex, "puts 'hello'.rindex(/[eo]/)", "4");
ruby_test!(test_rindex_offset, "puts 'hello'.rindex('l', 2)", "2");
ruby_test!(test_rindex_offset_negative, "puts 'hello'.rindex('l', -3)", "2");
ruby_test!(test_rindex_not_found, "puts 'hello'.rindex('z').nil?", "true");
ruby_test!(test_rindex_empty_string, "puts 'hello'.rindex('')", "5");
ruby_test!(test_rindex_empty_offset, "puts 'hello'.rindex('', 2)", "2");
