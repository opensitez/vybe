macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_product_basic,
    "puts [1, 2].product([3, 4]).inspect",
    "[[1, 3], [1, 4], [2, 3], [2, 4]]"
);
ruby_test!(
    test_product_multiple_arrays,
    "puts [1].product([2], [3]).inspect",
    "[[1, 2, 3]]"
);
ruby_test!(
    test_product_empty_arg,
    "puts [1, 2].product([]).inspect",
    "[]"
);
ruby_test!(
    test_product_empty_receiver,
    "puts [].product([1, 2]).inspect",
    "[]"
);
ruby_test!(
    test_product_no_args,
    "puts [1, 2].product.inspect",
    "[[1], [2]]"
);
ruby_test!(
    test_product_different_lengths,
    "puts [1].product([2, 3]).inspect",
    "[[1, 2], [1, 3]]"
);
ruby_test!(
    test_product_with_block,
    "acc = []; [1, 2].product([3, 4]) {|x, y| acc << \"#{x}#{y}\"}; puts acc.join('-')",
    "13-14-23-24"
);
ruby_test!(
    test_product_block_returns_self,
    "a = [1]; puts a.product([2]) {}.object_id == a.object_id",
    "true"
);
ruby_test!(
    test_product_mixed_types,
    "puts [1, 'a'].product([true]).inspect",
    "[[1, true], [\"a\", true]]"
);
ruby_test!(
    test_product_nested_arrays,
    "puts [[1]].product([[2]]).inspect",
    "[[[1], [2]]]"
);
ruby_test!(
    test_product_with_empty_block,
    "a = []; [1, 2].product([]) {|x| a << x}; puts a.length",
    "0"
);
