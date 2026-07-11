
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_slice_basic, "puts {a: 1, b: 2, c: 3}.slice(:a, :b).keys.sort.join('-')", "a-b");
ruby_test!(test_hash_slice_missing, "puts {a: 1}.slice(:b).empty?", "true");
ruby_test!(test_hash_slice_no_args, "puts {a: 1}.slice.empty?", "true");
ruby_test!(test_hash_except_basic, "puts {a: 1, b: 2, c: 3}.except(:a, :b).keys.sort.join('-')", "c");
ruby_test!(test_hash_except_missing, "puts {a: 1}.except(:b).keys.sort.join('-')", "a");
ruby_test!(test_hash_except_no_args, "puts {a: 1}.except.keys.sort.join('-')", "a");
