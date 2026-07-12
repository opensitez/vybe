macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_respond_to_missing_basic,
    "class A; def respond_to_missing?(m, include_private = false); m == :foo || super; end; end; puts A.new.respond_to?(:foo)",
    "true"
);
ruby_test!(
    test_respond_to_missing_false,
    "class A; def respond_to_missing?(m, include_private = false); m == :foo || super; end; end; puts A.new.respond_to?(:bar)",
    "false"
);
ruby_test!(
    test_respond_to_missing_method_call,
    "class A; def respond_to_missing?(m, inc); m == :foo; end; def method_missing(m, *args); 'foo'; end; end; puts A.new.method(:foo).call",
    "foo"
); // test if respond_to_missing? allows method(:foo)
