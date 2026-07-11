
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_zip_basic, "puts [1, 2].zip([3, 4]).inspect", "[[1, 3], [2, 4]]");
ruby_test!(test_zip_multiple_arrays, "puts [1, 2].zip([3, 4], [5, 6]).inspect", "[[1, 3, 5], [2, 4, 6]]");
ruby_test!(test_zip_different_lengths_short_arg, "puts [1, 2, 3].zip([4, 5]).inspect", "[[1, 4], [2, 5], [3, nil]]");
ruby_test!(test_zip_different_lengths_long_arg, "puts [1, 2].zip([3, 4, 5]).inspect", "[[1, 3], [2, 4]]"); // truncated to receiver length
ruby_test!(test_zip_empty_receiver, "puts [].zip([1, 2]).inspect", "[]");
ruby_test!(test_zip_empty_arg, "puts [1, 2].zip([]).inspect", "[[1, nil], [2, nil]]");
ruby_test!(test_zip_no_args, "puts [1, 2].zip.inspect", "[[1], [2]]");
ruby_test!(test_zip_with_block, "acc = []; [1, 2].zip([3, 4]) {|x, y| acc << x+y}; puts acc.join('-')", "4-6");
ruby_test!(test_zip_block_returns_nil, "puts [1, 2].zip([3, 4]) {}.nil?", "true");
ruby_test!(test_zip_non_array_arg, "puts [1, 2].zip(3..4).inspect", "[[1, 3], [2, 4]]"); // arguments are converted using to_ary or to_enum
ruby_test!(test_zip_nested_arrays, "puts [[1]].zip([[2]]).inspect", "[[[1], [2]]]");
