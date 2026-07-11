
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_each_slice_basic, "puts [1, 2, 3, 4, 5].each_slice(2).to_a.inspect", "[[1, 2], [3, 4], [5]]");
ruby_test!(test_each_slice_no_block, "puts [1, 2].each_slice(2).is_a?(Enumerator)", "true");
ruby_test!(test_each_slice_exact, "puts [1, 2, 3, 4].each_slice(2).to_a.inspect", "[[1, 2], [3, 4]]");
ruby_test!(test_each_slice_larger_than_length, "puts [1, 2].each_slice(5).to_a.inspect", "[[1, 2]]");
ruby_test!(test_each_slice_zero_error, "begin; [1].each_slice(0).to_a; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_each_slice_negative_error, "begin; [1].each_slice(-1).to_a; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_each_cons_basic, "puts [1, 2, 3, 4].each_cons(2).to_a.inspect", "[[1, 2], [2, 3], [3, 4]]");
ruby_test!(test_each_cons_no_block, "puts [1, 2].each_cons(2).is_a?(Enumerator)", "true");
ruby_test!(test_each_cons_larger_than_length, "puts [1, 2].each_cons(5).to_a.inspect", "[]");
ruby_test!(test_each_cons_zero_error, "begin; [1].each_cons(0).to_a; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_each_cons_negative_error, "begin; [1].each_cons(-1).to_a; rescue ArgumentError; puts 'err'; end", "err");
