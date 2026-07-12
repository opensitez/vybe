macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_array_search_find,
    "puts [1, 2, 3, 4].find { |x| x.even? }",
    "2"
);
ruby_test!(
    test_array_search_find_not_found,
    "puts [1, 3, 5].find { |x| x.even? }.nil?",
    "true"
);
ruby_test!(
    test_array_search_find_ifnone,
    "puts [1, 3].find(-> { 'none' }) { |x| x.even? }",
    "none"
);
ruby_test!(
    test_array_search_find_index,
    "puts [1, 2, 3, 4].find_index { |x| x.even? }",
    "1"
);
ruby_test!(
    test_array_search_find_index_value,
    "puts [1, 2, 3, 4].find_index(3)",
    "2"
);
ruby_test!(
    test_array_search_find_index_not_found,
    "puts [1, 2, 3].find_index(5).nil?",
    "true"
);
ruby_test!(
    test_array_search_rindex_value,
    "puts [1, 2, 2, 3].rindex(2)",
    "2"
);
ruby_test!(
    test_array_search_rindex_block,
    "puts [1, 2, 3, 4].rindex { |x| x.even? }",
    "3"
);
ruby_test!(
    test_array_search_include,
    "puts [1, 2, 3].include?(2)",
    "true"
);
ruby_test!(
    test_array_search_include_false,
    "puts [1, 2, 3].include?(4)",
    "false"
);
ruby_test!(
    test_array_search_any,
    "puts [1, 2, 3].any? { |x| x.even? }",
    "true"
);
ruby_test!(
    test_array_search_any_false,
    "puts [1, 3, 5].any? { |x| x.even? }",
    "false"
);
ruby_test!(
    test_array_search_all,
    "puts [2, 4, 6].all? { |x| x.even? }",
    "true"
);
ruby_test!(
    test_array_search_all_false,
    "puts [2, 4, 5].all? { |x| x.even? }",
    "false"
);
ruby_test!(
    test_array_search_none,
    "puts [1, 3, 5].none? { |x| x.even? }",
    "true"
);
ruby_test!(
    test_array_search_none_false,
    "puts [1, 2, 5].none? { |x| x.even? }",
    "false"
);
ruby_test!(
    test_array_search_one,
    "puts [1, 2, 5].one? { |x| x.even? }",
    "true"
);
ruby_test!(
    test_array_search_one_false,
    "puts [1, 2, 4].one? { |x| x.even? }",
    "false"
);
