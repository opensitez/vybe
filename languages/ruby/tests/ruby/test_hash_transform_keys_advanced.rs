macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_transform_keys_basic,
    "puts ({a: 1, b: 2}.transform_keys {|k| k.to_s}.keys.join('-'))",
    "a-b"
);
ruby_test!(
    test_transform_keys_no_block,
    "puts ({a: 1}.transform_keys.is_a?(Enumerator))",
    "true"
);
ruby_test!(
    test_transform_keys_duplicate_keys,
    "puts ({a: 1, b: 2}.transform_keys {|k| :c}.length)",
    "1"
); // later value wins
ruby_test!(
    test_transform_keys_duplicate_keys_value,
    "puts ({a: 1, b: 2}.transform_keys {|k| :c}[:c])",
    "2"
); // b's value overwrites a's
ruby_test!(
    test_transform_keys_returns_hash,
    "puts ({a: 1}.transform_keys {|k| k}.is_a?(Hash))",
    "true"
);
ruby_test!(
    test_transform_keys_does_not_mutate,
    "h = {a: 1}; h.transform_keys {|k| k.to_s}; puts h.keys[0].is_a?(Symbol)",
    "true"
);
ruby_test!(
    test_transform_keys_bang_mutates,
    "h = {a: 1, b: 2}; h.transform_keys! {|k| k.to_s}; puts h.keys.join('-')",
    "a-b"
);
ruby_test!(
    test_transform_keys_bang_returns_self,
    "h = {a: 1}; puts h.transform_keys! {|k| k}.object_id == h.object_id",
    "true"
);
ruby_test!(
    test_transform_keys_hash_arg,
    "puts ({a: 1, b: 2}.transform_keys({a: :c, b: :d}).keys.map(&:to_s).join('-'))",
    "c-d"
); // ruby 2.5+
ruby_test!(
    test_transform_keys_hash_arg_partial,
    "puts ({a: 1, b: 2}.transform_keys({a: :c}).keys.map(&:to_s).join('-'))",
    "c-b"
);
ruby_test!(
    test_transform_keys_hash_arg_and_block,
    "puts ({a: 1, b: 2}.transform_keys({a: :c}) {|k| k.to_s.upcase.to_sym}.keys.map(&:to_s).join('-'))",
    "c-B"
); // hash arg takes precedence over block
