macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_binding_eval_basic,
    "def foo; a = 1; binding; end; puts foo.eval('a + 1')",
    "2"
);
ruby_test!(
    test_binding_eval_mutation,
    "def foo; a = 1; b = binding; b.eval('a = 2'); a; end; puts foo",
    "2"
);
ruby_test!(
    test_binding_local_variables,
    "def foo; a = 1; b = 2; binding; end; puts foo.local_variables.sort.join('-')",
    "a-b"
);
ruby_test!(
    test_binding_local_variable_get,
    "def foo; a = 42; binding; end; puts foo.local_variable_get(:a)",
    "42"
);
ruby_test!(
    test_binding_local_variable_set,
    "def foo; a = 1; b = binding; b.local_variable_set(:a, 42); a; end; puts foo",
    "42"
);
ruby_test!(
    test_binding_local_variable_defined,
    "def foo; a = 1; binding; end; puts foo.local_variable_defined?(:a)",
    "true"
);
ruby_test!(
    test_binding_local_variable_defined_false,
    "def foo; a = 1; binding; end; puts foo.local_variable_defined?(:b)",
    "false"
);
ruby_test!(
    test_binding_receiver,
    "class C; def foo; binding; end; end; c = C.new; puts c.foo.receiver == c",
    "true"
);
