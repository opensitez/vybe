
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_methods_basic, "class A; def foo; end; end; puts A.new.methods.include?(:foo)", "true");
ruby_test!(test_methods_inherited, "class A; def foo; end; end; class B < A; end; puts B.new.methods.include?(:foo)", "true");
ruby_test!(test_methods_false_arg, "class A; def foo; end; end; class B < A; def bar; end; end; puts B.new.methods(false).include?(:foo)", "false"); // false means only own methods
ruby_test!(test_instance_methods_basic, "class A; def foo; end; end; puts A.instance_methods.include?(:foo)", "true");
ruby_test!(test_instance_methods_false_arg, "class A; def foo; end; end; class B < A; def bar; end; end; puts B.instance_methods(false).include?(:foo)", "false");
ruby_test!(test_public_methods_basic, "class A; def foo; end; end; puts A.new.public_methods.include?(:foo)", "true");
ruby_test!(test_private_methods_basic, "class A; private; def foo; end; end; puts A.new.private_methods.include?(:foo)", "true");
ruby_test!(test_protected_methods_basic, "class A; protected; def foo; end; end; puts A.new.protected_methods.include?(:foo)", "true");
ruby_test!(test_singleton_methods_basic, "obj = Object.new; def obj.foo; end; puts obj.singleton_methods.include?(:foo)", "true");
