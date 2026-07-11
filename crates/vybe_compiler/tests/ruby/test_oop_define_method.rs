
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_define_method_basic, "class A; define_method(:foo) { 'foo' }; end; puts A.new.foo", "foo");
ruby_test!(test_define_method_args, "class A; define_method(:foo) { |x| \"foo_#{x}\" }; end; puts A.new.foo(1)", "foo_1");
ruby_test!(test_define_method_closure, "class A; val = 'closure'; define_method(:foo) { val }; end; puts A.new.foo", "closure");
ruby_test!(test_define_method_private, "class A; define_method(:foo) { 'foo' }; private :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_define_singleton_method_basic, "obj = Object.new; obj.define_singleton_method(:foo) { 'foo' }; puts obj.foo", "foo");
