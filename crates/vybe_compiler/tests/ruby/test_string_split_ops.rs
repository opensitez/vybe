
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_split_basic, "puts 'a,b,c'.split(',').join('-')", "a-b-c");
ruby_test!(test_string_split_limit, "puts 'a,b,c'.split(',', 2).join('-')", "a-b,c");
ruby_test!(test_string_split_regexp, "puts 'a b,c'.split(/[, ]/).join('-')", "a-b-c");
ruby_test!(test_string_split_default, "puts ' a  b c '.split.join('-')", "a-b-c");
ruby_test!(test_string_partition, "puts 'hello'.partition('l').join('-')", "he-l-lo");
ruby_test!(test_string_rpartition, "puts 'hello'.rpartition('l').join('-')", "hel-l-o");
ruby_test!(test_string_partition_not_found, "puts 'hello'.partition('x').join('-')", "hello--");
ruby_test!(test_string_rpartition_not_found, "puts 'hello'.rpartition('x').join('-')", "--hello");
