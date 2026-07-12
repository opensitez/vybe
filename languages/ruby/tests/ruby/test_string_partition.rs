macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_string_partition_basic,
    "puts 'hello'.partition('l').join('-')",
    "he-l-lo"
);
ruby_test!(
    test_string_partition_regex,
    "puts 'hello'.partition(/[aeiou]/).join('-')",
    "h-e-llo"
);
ruby_test!(
    test_string_partition_not_found,
    "puts 'hello'.partition('x').join('-')",
    "hello--"
);
ruby_test!(
    test_string_rpartition_basic,
    "puts 'hello'.rpartition('l').join('-')",
    "hel-l-o"
);
ruby_test!(
    test_string_rpartition_regex,
    "puts 'hello'.rpartition(/[aeiou]/).join('-')",
    "hell-o-"
);
ruby_test!(
    test_string_rpartition_not_found,
    "puts 'hello'.rpartition('x').join('-')",
    "--hello"
);
ruby_test!(
    test_string_partition_empty,
    "puts 'hello'.partition('').join('-')",
    "-hello"
); // ruby partition on empty string returns ["", "", "hello"]
ruby_test!(
    test_string_rpartition_empty,
    "puts 'hello'.rpartition('').join('-')",
    "hello--"
); // rpartition on empty string returns ["hello", "", ""]
ruby_test!(
    test_string_partition_start,
    "puts 'hello'.partition('h').join('-')",
    "-h-ello"
);
ruby_test!(
    test_string_rpartition_end,
    "puts 'hello'.rpartition('o').join('-')",
    "hell-o-"
);
