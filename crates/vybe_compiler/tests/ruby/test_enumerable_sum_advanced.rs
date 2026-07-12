macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_sum_basic, "puts [1, 2, 3].sum", "6");
ruby_test!(test_sum_empty, "puts [].sum", "0");
ruby_test!(test_sum_init, "puts [1, 2, 3].sum(10)", "16");
ruby_test!(test_sum_empty_init, "puts [].sum(10)", "10");
ruby_test!(test_sum_block, "puts [1, 2, 3].sum {|x| x * 2}", "12");
ruby_test!(
    test_sum_block_init,
    "puts [1, 2, 3].sum(10) {|x| x * 2}",
    "22"
);
ruby_test!(
    test_sum_strings_error,
    "begin; ['a', 'b'].sum; rescue TypeError; puts 'err'; end",
    "err"
); // must supply init for strings
ruby_test!(test_sum_strings_init, "puts ['a', 'b'].sum('')", "ab");
ruby_test!(test_sum_floats, "puts [1.5, 2.0].sum", "3.5");
ruby_test!(test_sum_arrays, "puts [[1], [2]].sum([]).join('-')", "1-2");
