macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_min_by_basic,
    "puts ['apple', 'pear', 'fig'].min_by {|x| x.length}",
    "fig"
);
ruby_test!(
    test_min_by_count,
    "puts ['apple', 'pear', 'fig', 'a'].min_by(2) {|x| x.length}.join('-')",
    "a-fig"
);
ruby_test!(
    test_min_by_no_block,
    "puts [1].min_by.is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_max_by_basic,
    "puts ['apple', 'pear', 'fig'].max_by {|x| x.length}",
    "apple"
);
ruby_test!(
    test_max_by_count,
    "puts ['apple', 'pear', 'fig', 'a'].max_by(2) {|x| x.length}.join('-')",
    "apple-pear"
);
ruby_test!(
    test_max_by_no_block,
    "puts [1].max_by.is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_minmax_by_basic,
    "puts ['apple', 'pear', 'fig'].minmax_by {|x| x.length}.join('-')",
    "fig-apple"
);
ruby_test!(
    test_minmax_by_no_block,
    "puts [1].minmax_by.is_a?(Enumerator)",
    "true"
);
ruby_test!(test_min_by_empty, "puts [].min_by {|x| x}.nil?", "true");
ruby_test!(
    test_min_by_empty_count,
    "puts [].min_by(2) {|x| x}.length",
    "0"
);
ruby_test!(test_max_by_empty, "puts [].max_by {|x| x}.nil?", "true");
ruby_test!(
    test_minmax_by_empty,
    "puts [].minmax_by {|x| x}.inspect",
    "[nil, nil]"
);
