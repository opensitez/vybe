
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_binding_eval, "a = 1; puts binding.eval('a + 1')", "2");
ruby_test!(test_binding_local_variables, "a = 1; b = 2; puts binding.local_variables.sort.join('-')", "a-b");
ruby_test!(test_binding_local_variable_get, "a = 1; puts binding.local_variable_get(:a)", "1");
ruby_test!(test_binding_local_variable_set, "a = 1; binding.local_variable_set(:a, 2); puts a", "2");
ruby_test!(test_binding_local_variable_defined, "a = 1; puts binding.local_variable_defined?(:a)", "true");
ruby_test!(test_binding_local_variable_defined_false, "puts binding.local_variable_defined?(:b)", "false");
ruby_test!(test_binding_receiver, "class C; def foo; binding.receiver.class.name; end; end; puts C.new.foo", "C");
ruby_test!(test_binding_source_location, "puts binding.source_location.class.name", "Array");
