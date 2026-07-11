
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_multi_access, "a = [[1, 2], [3, 4]]; puts a[1][0]", "3");
ruby_test!(test_multi_flatten, "a = [[1, 2], [3, 4]]; puts a.flatten.join('-')", "1-2-3-4");
ruby_test!(test_multi_flatten_level, "a = [1, [2, [3, 4]]]; puts a.flatten(1).inspect", "[1, 2, [3, 4]]");
ruby_test!(test_multi_transpose, "a = [[1, 2], [3, 4]]; puts a.transpose.inspect", "[[1, 3], [2, 4]]");
ruby_test!(test_multi_transpose_error, "a = [[1, 2], [3]]; begin; a.transpose; rescue IndexError; puts 'err'; end", "err"); // must be same length
ruby_test!(test_multi_dig, "a = [[1, 2], [3, 4]]; puts a.dig(1, 1)", "4");
ruby_test!(test_multi_dig_missing, "a = [[1, 2], [3, 4]]; puts a.dig(2, 0).nil?", "true");
ruby_test!(test_multi_assoc, "a = [['a', 1], ['b', 2]]; puts a.assoc('b').join('-')", "b-2");
ruby_test!(test_multi_rassoc, "a = [['a', 1], ['b', 2]]; puts a.rassoc(1).join('-')", "a-1");
ruby_test!(test_multi_assoc_missing, "a = [['a', 1]]; puts a.assoc('c').nil?", "true");
ruby_test!(test_multi_rassoc_missing, "a = [['a', 1]]; puts a.rassoc(3).nil?", "true");
ruby_test!(test_multi_map_inner, "a = [[1, 2], [3, 4]]; puts a.map { |x| x.map { |y| y*2 } }.inspect", "[[2, 4], [6, 8]]");
