use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_partition_basic, "puts 'hello'.partition('l').join('-')", "he-l-lo");
ruby_test!(test_partition_not_found, "puts 'hello'.partition('z').join('-')", "hello--");
ruby_test!(test_partition_empty_search, "puts 'hello'.partition('').join('-')", "-h-ello");
ruby_test!(test_partition_regex, "puts 'hello'.partition(/l+/).join('-')", "he-ll-o");
ruby_test!(test_partition_regex_not_found, "puts 'hello'.partition(/z/).join('-')", "hello--");
ruby_test!(test_rpartition_basic, "puts 'hello'.rpartition('l').join('-')", "hel-l-o");
ruby_test!(test_rpartition_not_found, "puts 'hello'.rpartition('z').join('-')", "--hello");
ruby_test!(test_rpartition_empty_search, "puts 'hello'.rpartition('').join('-')", "hello--");
ruby_test!(test_rpartition_regex, "puts 'hello'.rpartition(/l/).join('-')", "hel-l-o");
ruby_test!(test_rpartition_regex_group, "puts 'hello'.rpartition(/(l)/).join('-')", "hel-l-o");
ruby_test!(test_partition_empty_string, "puts ''.partition('a').join('-')", "--");
ruby_test!(test_rpartition_empty_string, "puts ''.rpartition('a').join('-')", "--");
