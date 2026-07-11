
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_compare_eq, "puts ({a: 1, b: 2} == {b: 2, a: 1})", "true"); // order independent
ruby_test!(test_hash_compare_not_eq, "puts ({a: 1, b: 2} == {a: 1, b: 3})", "false");
ruby_test!(test_hash_compare_size, "puts ({a: 1} == {a: 1, b: 2})", "false");
ruby_test!(test_hash_compare_eql, "puts ({a: 1}.eql?({a: 1}))", "true"); // Hash#eql?
ruby_test!(test_hash_compare_lt, "puts ({a: 1} < {a: 1, b: 2})", "true"); // Subset
ruby_test!(test_hash_compare_lt_false, "puts ({a: 1} < {a: 1})", "false"); // Strict subset
ruby_test!(test_hash_compare_lte, "puts ({a: 1} <= {a: 1})", "true");
ruby_test!(test_hash_compare_gt, "puts ({a: 1, b: 2} > {a: 1})", "true"); // Superset
ruby_test!(test_hash_compare_gte, "puts ({a: 1} >= {a: 1})", "true");
ruby_test!(test_hash_compare_disjoint, "puts ({a: 1} < {b: 2}).nil?", "false"); // Wait, disjoint sets return false for <
ruby_test!(test_hash_compare_any_basic, "puts ({a: 1, b: 2}.any? { |k, v| v > 1 })", "true");
