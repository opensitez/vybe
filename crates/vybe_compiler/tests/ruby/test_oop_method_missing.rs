use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_method_missing_basic, "class A; def method_missing(m, *args); \"missing_#{m}\"; end; end; puts A.new.foo", "missing_foo");
ruby_test!(test_method_missing_args, "class A; def method_missing(m, *args); \"#{m}_#{args.join('-')}\"; end; end; puts A.new.foo(1, 2)", "foo_1-2");
ruby_test!(test_method_missing_super, "class A; def method_missing(m, *args); super; rescue NoMethodError; 'err'; end; end; puts A.new.foo", "err");
ruby_test!(test_method_missing_respond_to, "class A; def method_missing(m, *args); true; end; end; puts A.new.respond_to?(:foo)", "false"); // method_missing doesn't change respond_to? unless respond_to_missing? is defined
