use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_creation_literal, "puts ({a: 1, b: 2}.class.name)", "Hash");
ruby_test!(test_hash_creation_new, "puts Hash.new.class.name", "Hash");
ruby_test!(test_hash_creation_default_value, "h = Hash.new(42); puts h[:a]", "42");
ruby_test!(test_hash_creation_default_block, "h = Hash.new { |hash, key| hash[key] = key.to_s.upcase }; puts h[:a]", "A");
ruby_test!(test_hash_creation_brackets, "puts Hash['a', 1, 'b', 2]['b']", "2");
ruby_test!(test_hash_creation_brackets_array, "puts Hash[[['a', 1], ['b', 2]]]['b']", "2");
ruby_test!(test_hash_creation_invalid_brackets, "begin; Hash['a', 1, 'b']; rescue ArgumentError; puts 'err'; end", "err");
