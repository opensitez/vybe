
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_transform_values_basic, "puts ({a: 1, b: 2}.transform_values {|v| v * 2}.values.join('-'))", "2-4");
ruby_test!(test_transform_values_no_block, "puts ({a: 1}.transform_values.is_a?(Enumerator))", "true");
ruby_test!(test_transform_values_returns_hash, "puts ({a: 1}.transform_values {|v| v}.is_a?(Hash))", "true");
ruby_test!(test_transform_values_does_not_mutate, "h = {a: 1}; h.transform_values {|v| 2}; puts h[:a]", "1");
ruby_test!(test_transform_values_bang_mutates, "h = {a: 1, b: 2}; h.transform_values! {|v| v * 2}; puts h.values.join('-')", "2-4");
ruby_test!(test_transform_values_bang_returns_self, "h = {a: 1}; puts h.transform_values! {|v| v}.object_id == h.object_id", "true");
ruby_test!(test_transform_values_preserves_keys, "puts ({a: 1, b: 2}.transform_values {|v| 0}.keys.map(&:to_s).join('-'))", "a-b");
ruby_test!(test_transform_values_empty, "puts ({}).transform_values {|v| v}.length", "0");
ruby_test!(test_transform_values_nil_value, "puts ({a: nil}.transform_values {|v| 1}[:a])", "1");
ruby_test!(test_transform_values_preserves_default, "h = Hash.new('def'); h[:a] = 1; puts h.transform_values {|v| v}.default", "def");
