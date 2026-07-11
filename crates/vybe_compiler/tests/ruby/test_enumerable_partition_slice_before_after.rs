
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_partition_basic, "puts [1, 2, 3, 4].partition {|x| x % 2 == 0}.inspect", "[[2, 4], [1, 3]]");
ruby_test!(test_partition_all_true, "puts [2, 4].partition {|x| x % 2 == 0}.inspect", "[[2, 4], []]");
ruby_test!(test_partition_all_false, "puts [1, 3].partition {|x| x % 2 == 0}.inspect", "[[], [1, 3]]");
ruby_test!(test_partition_no_block, "puts [1].partition.is_a?(Enumerator)", "true");
ruby_test!(test_slice_before_pattern, "puts [1, 2, 3, 1, 2].slice_before(3).to_a.inspect", "[[1, 2], [3, 1, 2]]");
ruby_test!(test_slice_before_block, "puts [1, 2, 3, 4].slice_before {|x| x % 2 == 0}.to_a.inspect", "[[1], [2, 3], [4]]");
ruby_test!(test_slice_after_pattern, "puts [1, 2, 3, 1, 2].slice_after(3).to_a.inspect", "[[1, 2, 3], [1, 2]]");
ruby_test!(test_slice_after_block, "puts [1, 2, 3, 4].slice_after {|x| x % 2 == 0}.to_a.inspect", "[[1, 2], [3, 4]]");
ruby_test!(test_slice_when_basic, "puts [1, 2, 4, 5].slice_when {|i, j| i + 1 != j}.to_a.inspect", "[[1, 2], [4, 5]]");
ruby_test!(test_chunk_basic, "puts [1, 2, 3, 4, 5, 6].chunk {|x| x % 2 == 0}.to_a.inspect", "[[false, [1]], [true, [2]], [false, [3]], [true, [4]], [false, [5]], [true, [6]]]");
ruby_test!(test_chunk_while_basic, "puts [1, 2, 4, 5].chunk_while {|i, j| i + 1 == j}.to_a.inspect", "[[1, 2], [4, 5]]");
