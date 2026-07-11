
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_chunk, "puts [1, 1, 2, 2, 3].chunk { |n| n }.map { |k, v| \"#{k}:#{v.join(',')}\" }.join('-')", "1:1,1-2:2,2-3:3");
ruby_test!(test_enumerable_chunk_while, "puts [1, 2, 4, 9, 10, 11].chunk_while { |i, j| i + 1 == j }.map { |a| a.join(',') }.join('-')", "1,2-4-9,10,11");
ruby_test!(test_enumerable_slice_after, "puts [1, 2, 3, 4, 5].slice_after(&:even?).map { |a| a.join(',') }.join('-')", "1,2-3,4-5");
ruby_test!(test_enumerable_slice_before, "puts [1, 2, 3, 4, 5].slice_before(&:even?).map { |a| a.join(',') }.join('-')", "1-2,3-4,5");
ruby_test!(test_enumerable_slice_when, "puts [1, 2, 4, 9, 10, 11].slice_when { |i, j| i + 1 != j }.map { |a| a.join(',') }.join('-')", "1,2-4-9,10,11");
