use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_slice_basic, "puts ({a: 1, b: 2, c: 3}.slice(:a, :c).keys.join('-'))", "a-c");
ruby_test!(test_hash_slice_missing, "puts ({a: 1}.slice(:a, :b).keys.join('-'))", "a");
ruby_test!(test_hash_slice_empty, "puts ({a: 1}.slice().length)", "0");
ruby_test!(test_hash_slice_all, "puts ({a: 1, b: 2}.slice(:a, :b).keys.join('-'))", "a-b");
ruby_test!(test_hash_except_basic, "puts ({a: 1, b: 2, c: 3}.except(:b).keys.join('-'))", "a-c");
ruby_test!(test_hash_except_multiple, "puts ({a: 1, b: 2, c: 3}.except(:a, :c).keys.join('-'))", "b");
ruby_test!(test_hash_except_missing, "puts ({a: 1, b: 2}.except(:c).keys.join('-'))", "a-b");
ruby_test!(test_hash_except_empty, "puts ({a: 1, b: 2}.except().keys.join('-'))", "a-b");
ruby_test!(test_hash_except_all, "puts ({a: 1, b: 2}.except(:a, :b).length)", "0");
