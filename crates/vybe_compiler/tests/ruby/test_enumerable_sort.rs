macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerable_sort_basic,
    "puts [1, 5, 2].sort.join('-')",
    "1-2-5"
);
ruby_test!(
    test_enumerable_sort_block,
    "puts [1, 5, 2].sort { |a, b| b <=> a }.join('-')",
    "5-2-1"
);
ruby_test!(
    test_enumerable_sort_by,
    "puts %w[a abc ab].sort_by { |x| x.length }.join('-')",
    "a-ab-abc"
);
ruby_test!(
    test_enumerable_sort_by_no_block,
    "puts %w[a].sort_by.class.name",
    "Enumerator"
);
ruby_test!(
    test_enumerable_reverse_each,
    "acc = []; [1, 2, 3].reverse_each { |x| acc << x }; puts acc.join('-')",
    "3-2-1"
);
ruby_test!(
    test_enumerable_reverse_each_no_block,
    "puts [1].reverse_each.class.name",
    "Enumerator"
);
