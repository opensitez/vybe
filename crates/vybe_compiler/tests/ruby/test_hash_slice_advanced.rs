
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_slice_basic, "puts ({a: 1, b: 2, c: 3}.slice(:a, :c).keys.map(&:to_s).join('-'))", "a-c");
ruby_test!(test_slice_missing_keys, "puts ({a: 1}.slice(:a, :b).keys.map(&:to_s).join('-'))", "a");
ruby_test!(test_slice_no_args, "puts ({a: 1}.slice.length)", "0");
ruby_test!(test_slice_duplicate_keys, "puts ({a: 1}.slice(:a, :a).keys.map(&:to_s).join('-'))", "a");
ruby_test!(test_slice_returns_hash, "puts ({a: 1}.slice(:a).is_a?(Hash))", "true");
ruby_test!(test_slice_does_not_mutate, "h = {a: 1, b: 2}; h.slice(:a); puts h.length", "2");
ruby_test!(test_slice_ignores_hash_default, "h = Hash.new('def'); puts h.slice(:a).length", "0");
ruby_test!(test_slice_nil_value, "puts ({a: nil, b: 2}.slice(:a).values.inspect)", "[nil]");
ruby_test!(test_slice_empty_hash, "puts ({}).slice(:a).length", "0");
