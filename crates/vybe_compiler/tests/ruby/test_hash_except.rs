
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_except_basic, "puts ({a: 1, b: 2, c: 3}.except(:a, :c).keys.map(&:to_s).join('-'))", "b");
ruby_test!(test_except_missing_keys, "puts ({a: 1, b: 2}.except(:c).keys.map(&:to_s).join('-'))", "a-b");
ruby_test!(test_except_no_args, "puts ({a: 1}.except.keys.map(&:to_s).join('-'))", "a");
ruby_test!(test_except_duplicate_args, "puts ({a: 1, b: 2}.except(:a, :a).keys.map(&:to_s).join('-'))", "b");
ruby_test!(test_except_returns_hash, "puts ({a: 1}.except(:a).is_a?(Hash))", "true");
ruby_test!(test_except_does_not_mutate, "h = {a: 1, b: 2}; h.except(:a); puts h.length", "2");
ruby_test!(test_except_preserves_hash_default, "h = Hash.new('def'); puts h.except(:a).default", "def");
ruby_test!(test_except_empty_hash, "puts ({}).except(:a).length", "0");
ruby_test!(test_except_all_keys, "puts ({a: 1}.except(:a).length", "0");
