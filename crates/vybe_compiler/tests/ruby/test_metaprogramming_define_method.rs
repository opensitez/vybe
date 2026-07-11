
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_define_method_basic, "class A; define_method(:foo) { 'foo' }; end; puts A.new.foo", "foo");
ruby_test!(test_define_method_with_args, "class A; define_method(:foo) { |x| \"foo#{x}\" }; end; puts A.new.foo(1)", "foo1");
ruby_test!(test_define_singleton_method_basic, "o = Object.new; o.define_singleton_method(:foo) { 'foo' }; puts o.foo", "foo");
ruby_test!(test_remove_method_basic, "class A; def foo; 'foo'; end; remove_method :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_undef_method_basic, "class A; def foo; 'foo'; end; undef_method :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_undef_keyword_basic, "class A; def foo; 'foo'; end; undef foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "err");
