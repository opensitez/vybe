
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_transpose_basic, "puts [[1, 2], [3, 4]].transpose.inspect", "[[1, 3], [2, 4]]");
ruby_test!(test_transpose_non_square, "puts [[1, 2, 3], [4, 5, 6]].transpose.inspect", "[[1, 4], [2, 5], [3, 6]]");
ruby_test!(test_transpose_error_different_lengths, "begin; [[1, 2], [3]].transpose; rescue IndexError; puts 'err'; end", "err");
ruby_test!(test_transpose_empty_inner, "puts [[], []].transpose.inspect", "[]");
ruby_test!(test_transpose_empty_outer, "puts [].transpose.inspect", "[]");
ruby_test!(test_transpose_1d_error, "begin; [1, 2].transpose; rescue TypeError; puts 'err'; end", "err"); // elements must be arrays
ruby_test!(test_transpose_3d, "puts [[[1, 2]], [[3, 4]]].transpose.inspect", "[[[1, 2], [3, 4]]]");
ruby_test!(test_transpose_mixed_types_inner, "puts [[1, 'a'], [2, 'b']].transpose.inspect", "[[1, 2], [\"a\", \"b\"]]");
ruby_test!(test_transpose_single_row, "puts [[1, 2, 3]].transpose.inspect", "[[1], [2], [3]]");
ruby_test!(test_transpose_single_column, "puts [[1], [2], [3]].transpose.inspect", "[[1, 2, 3]]");
