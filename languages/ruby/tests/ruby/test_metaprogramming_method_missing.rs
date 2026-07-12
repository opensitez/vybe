macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_method_missing_basic,
    "class A; def method_missing(m, *args); \"missing #{m}\"; end; end; puts A.new.foo",
    "missing foo"
);
ruby_test!(
    test_method_missing_with_args,
    "class A; def method_missing(m, *args); \"missing #{m} #{args.join('-')}\"; end; end; puts A.new.foo(1, 2)",
    "missing foo 1-2"
);
ruby_test!(
    test_method_missing_block,
    "class A; def method_missing(m, *args, &block); \"missing #{m} #{block.call}\"; end; end; puts A.new.foo { 'block' }",
    "missing foo block"
);
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
    test_method_missing_super,
    "class A; def method_missing(m, *args); super; rescue NoMethodError; 'err'; end; end; puts A.new.foo",
    "err"
);
