
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_transform_keys, "puts {a: 1, b: 2}.transform_keys { |k| k.to_s.upcase }.keys.sort.join('-')", "A-B");
ruby_test!(test_hash_transform_keys_bang, "h = {a: 1, b: 2}; h.transform_keys! { |k| k.to_s.upcase }; puts h.keys.sort.join('-')", "A-B");
ruby_test!(test_hash_transform_values, "puts {a: 1, b: 2}.transform_values { |v| v * 2 }.values.sort.join('-')", "2-4");
ruby_test!(test_hash_transform_values_bang, "h = {a: 1, b: 2}; h.transform_values! { |v| v * 2 }; puts h.values.sort.join('-')", "2-4");
ruby_test!(test_hash_transform_keys_no_block, "puts {a: 1}.transform_keys.class.name", "Enumerator");
ruby_test!(test_hash_transform_values_no_block, "puts {a: 1}.transform_values.class.name", "Enumerator");
