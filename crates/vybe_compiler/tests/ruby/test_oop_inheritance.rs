
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_inheritance_basic, "class A; def foo; 'A'; end; end; class B < A; end; puts B.new.foo", "A");
ruby_test!(test_inheritance_override, "class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; end; puts B.new.foo", "B");
ruby_test!(test_inheritance_super_basic, "class A; def foo; 'A'; end; end; class B < A; def foo; super + 'B'; end; end; puts B.new.foo", "AB");
ruby_test!(test_inheritance_super_args, "class A; def foo(x); \"A#{x}\"; end; end; class B < A; def foo(x); super(x) + 'B'; end; end; puts B.new.foo(1)", "A1B");
ruby_test!(test_inheritance_super_implicit_args, "class A; def foo(x); \"A#{x}\"; end; end; class B < A; def foo(x); super + 'B'; end; end; puts B.new.foo(1)", "A1B"); // super without args passes args forward
ruby_test!(test_inheritance_superclass_method, "class A; end; class B < A; end; puts B.superclass == A", "true");
