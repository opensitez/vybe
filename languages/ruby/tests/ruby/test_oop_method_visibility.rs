macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_visibility_public_basic,
    "class A; def foo; 'foo'; end; end; puts A.new.foo",
    "foo"
);
ruby_test!(
    test_visibility_private_error,
    "class A; private; def foo; 'foo'; end; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_visibility_private_implicit_receiver,
    "class A; private; def foo; 'foo'; end; public; def bar; foo; end; end; puts A.new.bar",
    "foo"
);
ruby_test!(
    test_visibility_protected_error,
    "class A; protected; def foo; 'foo'; end; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_visibility_protected_same_class,
    "class A; protected; def foo; 'foo'; end; public; def bar(other); other.foo; end; end; puts A.new.bar(A.new)",
    "foo"
);
ruby_test!(
    test_visibility_protected_subclass,
    "class A; protected; def foo; 'foo'; end; end; class B < A; def bar(other); other.foo; end; end; puts B.new.bar(A.new)",
    "foo"
);
ruby_test!(
    test_visibility_method_calls,
    "class A; def foo; 'f'; end; private :foo; public :foo; protected :foo; end; puts A.new.respond_to?(:foo)",
    "false"
); // protected isn't public, respond_to? defaults to public
