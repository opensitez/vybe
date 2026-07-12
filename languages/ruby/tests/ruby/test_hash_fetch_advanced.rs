macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_fetch_basic, "puts ({a: 1, b: 2}.fetch(:a))", "1");
ruby_test!(
    test_fetch_missing,
    "begin; {}.fetch(:a); rescue KeyError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_fetch_default_value,
    "puts ({}.fetch(:a, 'def'))",
    "def"
);
ruby_test!(
    test_fetch_default_block,
    "puts ({}.fetch(:a) {|k| \"def#{k}\"})",
    "defa"
);
ruby_test!(
    test_fetch_block_precedence,
    "puts ({}.fetch(:a, 'def') {|k| 'blk'})",
    "blk"
); // block takes precedence
ruby_test!(
    test_fetch_ignores_hash_default,
    "h = Hash.new('hdef'); puts h.fetch(:a, 'fdef')",
    "fdef"
);
ruby_test!(
    test_fetch_ignores_hash_default_missing,
    "h = Hash.new('hdef'); begin; h.fetch(:a); rescue KeyError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_fetch_nil_value,
    "puts ({a: nil}.fetch(:a).nil?)",
    "true"
);
ruby_test!(
    test_fetch_nil_value_ignores_default,
    "puts ({a: nil}.fetch(:a, 'def').nil?)",
    "true"
); // exists, so default is ignored
ruby_test!(test_fetch_object_key, "puts ({[1] => 2}.fetch([1]))", "2");
