
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
ruby_test!(test_zip_different_lengths_shorter_arg, "puts [1, 2].zip([3]).inspect", "[[1, 3], [2, nil]]");
ruby_test!(test_zip_different_lengths_longer_arg, "puts [1, 2].zip([3, 4, 5]).inspect", "[[1, 3], [2, 4]]");
ruby_test!(test_zip_with_block, "acc = []; [1, 2].zip([3, 4]) {|x, y| acc << x + y}; puts acc.join('-')", "4-6");
ruby_test!(test_zip_no_args, "puts [1, 2].zip.inspect", "[[1], [2]]");
ruby_test!(test_zip_empty, "puts [].zip([1, 2]).inspect", "[]");
ruby_test!(test_zip_non_array_arg, "class A; def to_ary; [3, 4]; end; end; puts [1, 2].zip(A.new).inspect", "[[1, 3], [2, 4]]");
ruby_test!(test_zip_non_array_arg_enumerator, "puts [1, 2].zip(3..4).inspect", "[[1, 3], [2, 4]]"); // ruby converts arg to array if it responds to each
ruby_test!(test_zip_hash, "puts ({a: 1}.zip({b: 2}).inspect)", "[[[:a, 1], [:b, 2]]]");
