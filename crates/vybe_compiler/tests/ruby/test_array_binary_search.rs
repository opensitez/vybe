macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_array_bsearch,
    "puts [1, 2, 3, 4, 5].bsearch { |x| x >= 3 }",
    "3"
);
ruby_test!(
    test_array_bsearch_not_found,
    "puts [1, 2, 3].bsearch { |x| x > 5 }.nil?",
    "true"
);
ruby_test!(
    test_array_bsearch_index,
    "puts [1, 2, 3, 4, 5].bsearch_index { |x| x >= 3 }",
    "2"
);
ruby_test!(
    test_array_bsearch_index_not_found,
    "puts [1, 2, 3].bsearch_index { |x| x > 5 }.nil?",
    "true"
);
