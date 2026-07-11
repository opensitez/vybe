
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_env_get, "ENV['VYBE_TEST_ENV'] = 'hello'; puts ENV['VYBE_TEST_ENV']", "hello");
ruby_test!(test_env_set, "ENV['VYBE_TEST_ENV'] = 'world'; puts ENV['VYBE_TEST_ENV']", "world");
ruby_test!(test_env_keys, "ENV['VYBE_TEST_ENV'] = '1'; puts ENV.keys.include?('VYBE_TEST_ENV')", "true");
ruby_test!(test_env_values, "ENV['VYBE_TEST_ENV'] = 'val123'; puts ENV.values.include?('val123')", "true");
ruby_test!(test_env_delete, "ENV['VYBE_TEST_ENV'] = '1'; ENV.delete('VYBE_TEST_ENV'); puts ENV['VYBE_TEST_ENV'].nil?", "true");
ruby_test!(test_env_has_key, "ENV['VYBE_TEST_ENV'] = '1'; puts ENV.has_key?('VYBE_TEST_ENV')", "true");
ruby_test!(test_env_fetch, "ENV['VYBE_TEST_ENV'] = 'fetch'; puts ENV.fetch('VYBE_TEST_ENV')", "fetch");
ruby_test!(test_env_fetch_default, "puts ENV.fetch('NON_EXISTENT', 'default')", "default");
