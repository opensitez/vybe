macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_inject_basic_block,
    "puts [1, 2, 3].inject(0) {|sum, n| sum + n}",
    "6"
);
ruby_test!(
    test_inject_no_init_block,
    "puts [1, 2, 3].inject {|sum, n| sum + n}",
    "6"
); // uses first element as initial
ruby_test!(test_inject_symbol, "puts [1, 2, 3].inject(:+)", "6");
ruby_test!(
    test_inject_init_symbol,
    "puts [1, 2, 3].inject(10, :+)",
    "16"
);
ruby_test!(
    test_reduce_alias_basic,
    "puts [1, 2, 3].reduce(0) {|sum, n| sum + n}",
    "6"
);
ruby_test!(test_reduce_alias_symbol, "puts [1, 2, 3].reduce(:*)", "6");
ruby_test!(
    test_inject_empty_no_init,
    "puts [].inject {|sum, n| sum + n}.nil?",
    "true"
);
ruby_test!(
    test_inject_empty_with_init,
    "puts [].inject(5) {|sum, n| sum + n}",
    "5"
);
ruby_test!(
    test_inject_empty_symbol_no_init,
    "puts [].inject(:+).nil?",
    "true"
);
ruby_test!(
    test_inject_empty_symbol_with_init,
    "puts [].inject(5, :+)",
    "5"
);
ruby_test!(
    test_inject_hash,
    "puts ({a: 1, b: 2}.inject(0) {|sum, kv| sum + kv[1]})",
    "3"
);
