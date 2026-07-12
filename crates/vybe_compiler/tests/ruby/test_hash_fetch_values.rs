macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_hash_fetch_values_basic,
    "puts {a: 1, b: 2}.fetch_values(:a, :b).join('-')",
    "1-2"
);
ruby_test!(
    test_hash_fetch_values_missing_raise,
    "begin; {a: 1}.fetch_values(:a, :b); rescue KeyError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_hash_fetch_values_missing_block,
    "puts {a: 1}.fetch_values(:a, :b) { |k| k.to_s.upcase }.join('-')",
    "1-B"
);
ruby_test!(
    test_hash_values_at,
    "puts {a: 1, b: 2}.values_at(:a, :c).map(&:to_s).join('-')",
    "1-"
);
ruby_test!(
    test_hash_values_at_multiple,
    "puts {a: 1, b: 2}.values_at(:a, :a, :b).join('-')",
    "1-1-2"
);
