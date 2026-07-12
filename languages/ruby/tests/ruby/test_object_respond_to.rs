macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_respond_to_basic,
    "class A; def foo; end; end; puts A.new.respond_to?(:foo)",
    "true"
);
ruby_test!(
    test_respond_to_missing,
    "class A; def foo; end; end; puts A.new.respond_to?(:bar)",
    "false"
);
ruby_test!(
    test_respond_to_private,
    "class A; private; def foo; end; end; puts A.new.respond_to?(:foo)",
    "false"
);
ruby_test!(
    test_respond_to_private_include_private,
    "class A; private; def foo; end; end; puts A.new.respond_to?(:foo, true)",
    "true"
);
ruby_test!(
    test_respond_to_string_name,
    "class A; def foo; end; end; puts A.new.respond_to?('foo')",
    "true"
);
