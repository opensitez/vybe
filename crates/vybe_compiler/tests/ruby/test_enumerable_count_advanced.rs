macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_count_basic, "puts [1, 2, 3].count", "3");
ruby_test!(test_count_empty, "puts [].count", "0");
ruby_test!(test_count_arg, "puts [1, 2, 1, 3].count(1)", "2");
ruby_test!(test_count_arg_missing, "puts [1, 2, 3].count(4)", "0");
ruby_test!(
    test_count_block,
    "puts [1, 2, 3, 4].count {|x| x % 2 == 0}",
    "2"
);
ruby_test!(
    test_count_block_false,
    "puts [1, 3, 5].count {|x| x % 2 == 0}",
    "0"
);
ruby_test!(
    test_count_arg_and_block,
    "puts [1, 2].count(1) {|x| x > 0}",
    "1"
); // arg is favored, block generates warning in real ruby but ignores block result
ruby_test!(test_count_hash, "puts ({a: 1, b: 2}.count)", "2");
ruby_test!(
    test_count_hash_block,
    "puts ({a: 1, b: 2}.count {|k, v| v > 1})",
    "1"
);
ruby_test!(test_count_nil, "puts [1, nil, 2, nil].count(nil)", "2");
