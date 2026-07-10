use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_object_eq_basic, "o = Object.new; puts o == o", "true");
ruby_test!(test_object_eq_diff, "puts Object.new == Object.new", "false");
ruby_test!(test_object_eql_basic, "o = Object.new; puts o.eql?(o)", "true");
ruby_test!(test_object_eql_diff, "puts Object.new.eql?(Object.new)", "false");
ruby_test!(test_object_equal_basic, "o = Object.new; puts o.equal?(o)", "true");
ruby_test!(test_object_equal_diff, "puts Object.new.equal?(Object.new)", "false");
ruby_test!(test_object_equal_override, "class A; def equal?(other); true; end; end; puts A.new.equal?(Object.new)", "true"); // equal? should generally not be overridden, but it can be
ruby_test!(test_object_triple_eq_basic, "o = Object.new; puts (o === o)", "true");
ruby_test!(test_object_triple_eq_diff, "puts (Object.new === Object.new)", "false");
ruby_test!(test_object_not_eq_basic, "o = Object.new; puts o != o", "false");
ruby_test!(test_object_not_eq_diff, "puts Object.new != Object.new", "true");
