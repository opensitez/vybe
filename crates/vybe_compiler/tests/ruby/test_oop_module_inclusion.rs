use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_module_include_basic, "module M; def foo; 'M'; end; end; class A; include M; end; puts A.new.foo", "M");
ruby_test!(test_module_include_override, "module M; def foo; 'M'; end; end; class A; include M; def foo; 'A'; end; end; puts A.new.foo", "A");
ruby_test!(test_module_include_multiple, "module M1; def foo; 'M1'; end; end; module M2; def foo; 'M2'; end; end; class A; include M1; include M2; end; puts A.new.foo", "M2"); // Last included is first in ancestors after class
ruby_test!(test_module_include_super, "module M; def foo; 'M'; end; end; class A; include M; def foo; super + 'A'; end; end; puts A.new.foo", "MA");
ruby_test!(test_module_included_modules, "module M; end; class A; include M; end; puts A.included_modules.include?(M)", "true");
ruby_test!(test_module_ancestors, "module M; end; class A; include M; end; puts A.ancestors[0..1].join('-')", "A-M");
