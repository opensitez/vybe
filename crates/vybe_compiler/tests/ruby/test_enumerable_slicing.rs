use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_chunk, "puts [1, 2, 2, 3, 4, 4].chunk { |x| x.even? }.map { |even, arr| \"#{even}:#{arr.join(',')}\" }.join('-')", "false:1-true:2,2-false:3-true:4,4");
ruby_test!(test_enumerable_chunk_while, "puts [1, 2, 4, 5, 7].chunk_while { |i, j| i + 1 == j }.map { |arr| arr.join(',') }.join('-')", "1,2-4,5-7");
ruby_test!(test_enumerable_slice_after, "puts [1, 2, 3, 4].slice_after { |x| x.even? }.map { |arr| arr.join(',') }.join('-')", "1,2-3,4");
ruby_test!(test_enumerable_slice_after_pattern, "puts ['a', 'b', 'c'].slice_after(/b/).map { |arr| arr.join(',') }.join('-')", "a,b-c");
ruby_test!(test_enumerable_slice_before, "puts [1, 2, 3, 4].slice_before { |x| x.even? }.map { |arr| arr.join(',') }.join('-')", "1-2,3-4");
ruby_test!(test_enumerable_slice_before_pattern, "puts ['a', 'b', 'c'].slice_before(/b/).map { |arr| arr.join(',') }.join('-')", "a-b,c");
ruby_test!(test_enumerable_slice_when, "puts [1, 2, 4, 5, 7].slice_when { |i, j| i + 1 != j }.map { |arr| arr.join(',') }.join('-')", "1,2-4,5-7");
