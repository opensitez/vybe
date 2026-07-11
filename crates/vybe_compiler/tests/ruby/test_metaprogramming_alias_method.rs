
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_alias_method_basic, "class A; def foo; 'foo'; end; alias_method :bar, :foo; end; puts A.new.bar", "foo");
ruby_test!(test_alias_method_string_names, "class A; def foo; 'foo'; end; alias_method 'bar', 'foo'; end; puts A.new.bar", "foo");
ruby_test!(test_alias_method_private, "class A; private; def foo; 'foo'; end; public; alias_method :bar, :foo; end; puts A.new.bar", "foo"); // aliases copy visibility or can be made public? In Ruby, alias_method copies the visibility of the original method, wait, actually no, alias_method keeps the current visibility setting, whereas `alias` keeps the original visibility?
// Let's test a simple alias
ruby_test!(test_alias_method_override, "class A; def foo; 'foo1'; end; alias_method :bar, :foo; def foo; 'foo2'; end; end; puts A.new.bar", "foo1");
