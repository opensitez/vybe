macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_env_fetch,
    "ENV['FOO'] = 'bar'; puts ENV.fetch('FOO')",
    "bar"
);
ruby_test!(
    test_env_fetch_default,
    "puts ENV.fetch('MISSING', 'def')",
    "def"
);
ruby_test!(
    test_env_fetch_block,
    "puts ENV.fetch('MISSING') { |k| k.upcase }",
    "MISSING"
);
ruby_test!(
    test_env_store,
    "ENV.store('FOO', 'baz'); puts ENV['FOO']",
    "baz"
);
ruby_test!(
    test_env_keys,
    "ENV['FOO'] = '1'; puts ENV.keys.include?('FOO').to_s",
    "true"
);
ruby_test!(
    test_env_values,
    "ENV['FOO'] = 'bar'; puts ENV.values.include?('bar').to_s",
    "true"
);
ruby_test!(
    test_env_each,
    "ENV['FOO'] = 'bar'; found = false; ENV.each { |k, v| found = true if k == 'FOO' && v == 'bar' }; puts found",
    "true"
);
ruby_test!(
    test_env_delete,
    "ENV['FOO'] = 'bar'; ENV.delete('FOO'); puts ENV.has_key?('FOO')",
    "false"
);
ruby_test!(
    test_env_has_key,
    "ENV['FOO'] = '1'; puts ENV.has_key?('FOO')",
    "true"
);
ruby_test!(
    test_env_has_value,
    "ENV['FOO'] = 'bar'; puts ENV.has_value?('bar')",
    "true"
);
ruby_test!(
    test_env_to_h,
    "ENV['FOO'] = 'bar'; puts ENV.to_h['FOO']",
    "bar"
);
ruby_test!(
    test_env_clear,
    "ENV['FOO'] = '1'; ENV.clear; puts ENV.empty?",
    "true"
);
