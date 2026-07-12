macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_object_instance_eval,
    "obj = Object.new; obj.instance_eval('def foo; 1; end'); puts obj.foo",
    "1"
);
ruby_test!(
    test_object_instance_exec,
    "obj = Object.new; obj.instance_exec(42) { |x| define_singleton_method(:foo) { x } }; puts obj.foo",
    "42"
);
ruby_test!(
    test_object_instance_variables,
    "class C; def initialize; @a = 1; @b = 2; end; end; puts C.new.instance_variables.sort.join('-')",
    "@a-@b"
);
ruby_test!(
    test_object_instance_variable_get,
    "class C; def initialize; @a = 1; end; end; puts C.new.instance_variable_get(:@a)",
    "1"
);
ruby_test!(
    test_object_instance_variable_set,
    "obj = Object.new; obj.instance_variable_set(:@a, 1); puts obj.instance_variable_get(:@a)",
    "1"
);
ruby_test!(
    test_object_instance_variable_defined,
    "obj = Object.new; obj.instance_variable_set(:@a, 1); puts obj.instance_variable_defined?(:@a)",
    "true"
);
ruby_test!(
    test_object_remove_instance_variable,
    "obj = Object.new; obj.instance_variable_set(:@a, 1); obj.send(:remove_instance_variable, :@a); puts obj.instance_variable_defined?(:@a)",
    "false"
);
ruby_test!(
    test_object_singleton_methods,
    "obj = Object.new; def obj.foo; 1; end; def obj.bar; 2; end; puts obj.singleton_methods.sort.join('-')",
    "bar-foo"
);
ruby_test!(
    test_object_methods,
    "puts Object.new.methods.include?(:to_s)",
    "true"
);
ruby_test!(
    test_object_public_methods,
    "puts Object.new.public_methods.include?(:to_s)",
    "true"
);
ruby_test!(
    test_object_private_methods,
    "class C; private; def foo; end; end; puts C.new.private_methods.include?(:foo)",
    "true"
);
ruby_test!(
    test_object_protected_methods,
    "class C; protected; def foo; end; end; puts C.new.protected_methods.include?(:foo)",
    "true"
);
