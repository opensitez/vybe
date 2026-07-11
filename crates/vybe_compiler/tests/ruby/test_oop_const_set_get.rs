
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_const_set_basic, "class A; end; A.const_set(:C, 'C'); puts A::C", "C");
ruby_test!(test_const_set_string_name, "class A; end; A.const_set('C', 'C'); puts A::C", "C");
ruby_test!(test_const_get_basic, "class A; C = 'C'; end; puts A.const_get(:C)", "C");
ruby_test!(test_const_get_inherited, "class A; C = 'C'; end; class B < A; end; puts B.const_get(:C)", "C");
ruby_test!(test_const_get_false_arg, "class A; C = 'C'; end; class B < A; end; begin; B.const_get(:C, false); rescue NameError; puts 'err'; end", "err"); // false means don't inherit
ruby_test!(test_const_defined_basic, "class A; C = 'C'; end; puts A.const_defined?(:C)", "true");
ruby_test!(test_const_defined_inherited, "class A; C = 'C'; end; class B < A; end; puts B.const_defined?(:C)", "true");
ruby_test!(test_const_defined_false_arg, "class A; C = 'C'; end; class B < A; end; puts B.const_defined?(:C, false)", "false");
ruby_test!(test_constants_basic, "class A; C = 'C'; D = 'D'; end; puts A.constants.sort.join('-')", "C-D");
