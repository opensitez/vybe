use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_module_prepend_basic, "module M; def foo; 'M'; end; end; class A; prepend M; end; puts A.new.foo", "M");
ruby_test!(test_module_prepend_override, "module M; def foo; 'M'; end; end; class A; prepend M; def foo; 'A'; end; end; puts A.new.foo", "M"); // Prepend wins over class method
ruby_test!(test_module_prepend_super, "module M; def foo; super + 'M'; end; end; class A; prepend M; def foo; 'A'; end; end; puts A.new.foo", "AM");
ruby_test!(test_module_prepend_multiple, "module M1; def foo; 'M1'; end; end; module M2; def foo; 'M2'; end; end; class A; prepend M1; prepend M2; end; puts A.new.foo", "M2"); // Last prepended is first in ancestors
ruby_test!(test_module_prepend_ancestors, "module M; end; class A; prepend M; end; puts A.ancestors[0..1].join('-')", "M-A");
