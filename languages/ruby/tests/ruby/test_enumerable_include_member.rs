macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_include_basic, "puts [1, 2, 3].include?(2)", "true");
ruby_test!(test_include_missing, "puts [1, 2, 3].include?(4)", "false");
ruby_test!(
    test_include_type_mismatch,
    "puts [1, 2, 3].include?('2')",
    "false"
); // uses ==
ruby_test!(test_include_coercion, "puts [1.0, 2.0].include?(1)", "true"); // 1.0 == 1 is true
ruby_test!(test_member_alias, "puts [1, 2, 3].member?(2)", "true");
ruby_test!(test_include_empty, "puts [].include?(1)", "false");
ruby_test!(test_include_nil, "puts [1, nil, 2].include?(nil)", "true");
ruby_test!(
    test_include_hash_key,
    "puts ({a: 1}.include?([:a, 1]))",
    "false"
); // Hash#include? checks keys, not key/value entry arrays.
ruby_test!(
    test_member_hash_key,
    "puts ({a: 1}.member?([:a, 1]))",
    "false"
); // Hash#member? is also key membership.
ruby_test!(
    test_hash_include_overridden,
    "puts ({a: 1}.include?(:a))",
    "true"
);
