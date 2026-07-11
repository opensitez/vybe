
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_conversion_to_a, "h = {a: 1, b: 2}; puts h.to_a.map { |pair| pair.join(':') }.join('-')", "a:1-b:2");
ruby_test!(test_hash_conversion_to_h, "h = {a: 1}; puts h.to_h.equal?(h)", "true"); // to_h on Hash returns self
ruby_test!(test_hash_conversion_to_h_block, "h = {a: 1, b: 2}; puts h.to_h { |k, v| [k.to_s, v * 10] }['a']", "10");
ruby_test!(test_hash_conversion_flatten, "h = {a: 1, b: 2}; puts h.flatten.join('-')", "a-1-b-2");
ruby_test!(test_hash_conversion_flatten_level, "h = {a: [1, 2]}; puts h.flatten(1).map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "a-arr");
