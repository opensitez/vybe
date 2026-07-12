macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_hash_access_bracket, "h = {a: 1}; puts h[:a]", "1");
ruby_test!(
    test_hash_access_missing,
    "h = {a: 1}; puts h[:b].nil?",
    "true"
);
ruby_test!(
    test_hash_access_keys,
    "h = {a: 1, b: 2}; puts h.keys.join('-')",
    "a-b"
);
ruby_test!(
    test_hash_access_values,
    "h = {a: 1, b: 2}; puts h.values.join('-')",
    "1-2"
);
ruby_test!(
    test_hash_access_values_at,
    "h = {a: 1, b: 2, c: 3}; puts h.values_at(:a, :c).join('-')",
    "1-3"
);
ruby_test!(
    test_hash_access_has_key,
    "h = {a: 1}; puts h.has_key?(:a)",
    "true"
);
ruby_test!(
    test_hash_access_include,
    "h = {a: 1}; puts h.include?(:a)",
    "true"
);
ruby_test!(
    test_hash_access_key_question,
    "h = {a: 1}; puts h.key?(:a)",
    "true"
);
ruby_test!(
    test_hash_access_member,
    "h = {a: 1}; puts h.member?(:a)",
    "true"
);
ruby_test!(
    test_hash_access_has_value,
    "h = {a: 1}; puts h.has_value?(1)",
    "true"
);
ruby_test!(
    test_hash_access_value_question,
    "h = {a: 1}; puts h.value?(1)",
    "true"
);
ruby_test!(test_hash_access_empty, "puts {}.empty?", "true");
ruby_test!(test_hash_access_length, "puts {a: 1, b: 2}.length", "2");
ruby_test!(test_hash_access_size, "puts {a: 1, b: 2}.size", "2");
