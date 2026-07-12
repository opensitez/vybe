macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_enumerable_min_basic, "puts [1, 5, 2].min", "1");
ruby_test!(
    test_enumerable_min_block,
    "puts %w[a abc ab].min { |a, b| a.length <=> b.length }",
    "a"
);
ruby_test!(test_enumerable_max_basic, "puts [1, 5, 2].max", "5");
ruby_test!(
    test_enumerable_max_block,
    "puts %w[a abc ab].max { |a, b| a.length <=> b.length }",
    "abc"
);
ruby_test!(
    test_enumerable_minmax_basic,
    "puts [1, 5, 2].minmax.join('-')",
    "1-5"
);
ruby_test!(
    test_enumerable_minmax_block,
    "puts %w[a abc ab].minmax { |a, b| a.length <=> b.length }.join('-')",
    "a-abc"
);
ruby_test!(
    test_enumerable_min_by,
    "puts %w[a abc ab].min_by { |x| x.length }",
    "a"
);
ruby_test!(
    test_enumerable_max_by,
    "puts %w[a abc ab].max_by { |x| x.length }",
    "abc"
);
ruby_test!(
    test_enumerable_minmax_by,
    "puts %w[a abc ab].minmax_by { |x| x.length }.join('-')",
    "a-abc"
);
