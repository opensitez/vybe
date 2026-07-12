macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_array_bsearch_find,
    "puts [1, 2, 4, 8, 16].bsearch { |x| x >= 4 }",
    "4"
);
ruby_test!(
    test_array_bsearch_find_missing,
    "puts [1, 2, 4, 8].bsearch { |x| x >= 10 }.nil?",
    "true"
);
ruby_test!(
    test_array_bsearch_compare,
    "puts [1, 2, 4, 8, 16].bsearch { |x| 4 <=> x }",
    "4"
);
ruby_test!(
    test_array_bsearch_compare_missing,
    "puts [1, 2, 4, 8, 16].bsearch { |x| 5 <=> x }.nil?",
    "true"
);
ruby_test!(
    test_array_bsearch_index_find,
    "puts [1, 2, 4, 8, 16].bsearch_index { |x| x >= 4 }",
    "2"
);
ruby_test!(
    test_array_bsearch_index_missing,
    "puts [1, 2, 4, 8].bsearch_index { |x| x >= 10 }.nil?",
    "true"
);
ruby_test!(
    test_array_bsearch_index_compare,
    "puts [1, 2, 4, 8, 16].bsearch_index { |x| 4 <=> x }",
    "2"
);
ruby_test!(
    test_array_bsearch_index_compare_missing,
    "puts [1, 2, 4, 8, 16].bsearch_index { |x| 5 <=> x }.nil?",
    "true"
);
ruby_test!(
    test_array_bsearch_enumerator,
    "puts [1, 2, 3].bsearch.class.name",
    "Enumerator"
);
ruby_test!(
    test_array_bsearch_index_enumerator,
    "puts [1, 2, 3].bsearch_index.class.name",
    "Enumerator"
);
