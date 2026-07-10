use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_eql_basic, "puts {a: 1}.eql?({a: 1})", "true");
ruby_test!(test_eql_different_order, "puts {a: 1, b: 2}.eql?({b: 2, a: 1})", "true");
ruby_test!(test_eql_different_values, "puts {a: 1}.eql?({a: 2})", "false");
ruby_test!(test_eql_different_keys, "puts {a: 1}.eql?({b: 1})", "false");
ruby_test!(test_eql_different_length, "puts {a: 1}.eql?({a: 1, b: 2})", "false");
ruby_test!(test_eql_type_mismatch, "puts {a: 1}.eql?([1])", "false");
ruby_test!(test_eql_nested, "puts {a: {b: 1}}.eql?({a: {b: 1}})", "true");
ruby_test!(test_eql_default_ignored, "h1 = Hash.new(1); h2 = Hash.new(2); puts h1.eql?(h2)", "true"); // default value is ignored in comparison!
ruby_test!(test_equal_operator_basic, "puts ({a: 1} == {a: 1})", "true");
ruby_test!(test_equal_operator_different_order, "puts ({a: 1, b: 2} == {b: 2, a: 1})", "true");
ruby_test!(test_equal_operator_type_mismatch, "puts ({a: 1} == [1])", "false");
ruby_test!(test_equal_operator_nested, "puts ({a: {b: 1}} == {a: {b: 1}})", "true");
ruby_test!(test_equal_operator_value_coercion, "puts ({a: 1.0} == {a: 1})", "true"); // == uses == for values
ruby_test!(test_eql_value_no_coercion, "puts {a: 1.0}.eql?({a: 1})", "false"); // eql? uses eql? for values (usually)
