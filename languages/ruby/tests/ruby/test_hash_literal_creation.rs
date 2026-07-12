macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_literal_basic,
    "puts ({'a' => 1, 'b' => 2}.length)",
    "2"
);
ruby_test!(test_literal_symbol_keys, "puts ({a: 1, b: 2}.length)", "2");
ruby_test!(
    test_literal_mixed_keys,
    "puts ({'a' => 1, b: 2}.length)",
    "2"
);
ruby_test!(test_literal_empty, "puts ({}).length", "0");
ruby_test!(
    test_literal_string_interpolation_key,
    "x = 'c'; puts ({ \"#{x}\" => 3 }['c'])",
    "3"
);
ruby_test!(
    test_literal_symbol_interpolation_key,
    "x = 'c'; puts ({ \"#{x}\": 3 }[:c])",
    "3"
); // syntax added in Ruby 2.2
ruby_test!(
    test_literal_integer_keys,
    "puts ({1 => 'a', 2 => 'b'}[1])",
    "a"
);
ruby_test!(test_literal_float_keys, "puts ({1.5 => 'a'}[1.5])", "a");
ruby_test!(
    test_literal_array_keys,
    "puts ({[1, 2] => 'a'}[[1, 2]])",
    "a"
);
ruby_test!(
    test_literal_hash_keys,
    "puts ({{a: 1} => 'b'}[{a: 1}])",
    "b"
);
ruby_test!(test_literal_duplicate_keys, "puts ({a: 1, a: 2}[:a])", "2"); // later key overwrites
