use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_singleton_methods_basic, "o = Object.new; def o.foo; end; puts o.singleton_methods.include?(:foo)", "true");
ruby_test!(test_singleton_methods_false_arg, "class A; def self.foo; end; end; class B < A; def self.bar; end; end; puts B.singleton_methods(false).include?(:foo)", "false");
ruby_test!(test_singleton_method_added, "class A; @acc = []; def self.singleton_method_added(m); @acc << m unless m == :singleton_method_added; end; def self.foo; end; def self.acc; @acc; end; end; puts A.acc.include?(:foo)", "true");
ruby_test!(test_singleton_class_eval, "o = Object.new; o.singleton_class.class_eval { def foo; 'foo'; end }; puts o.foo", "foo");
