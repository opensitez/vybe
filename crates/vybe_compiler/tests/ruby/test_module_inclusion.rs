use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_module_inclusion_basic, "module M; def foo; 1; end; end; class C; include M; end; puts C.new.foo", "1");
ruby_test!(test_module_inclusion_override, "module M; def foo; 1; end; end; class C; include M; def foo; 2; end; end; puts C.new.foo", "2");
ruby_test!(test_module_inclusion_super, "module M; def foo; 1; end; end; class C; include M; def foo; super + 1; end; end; puts C.new.foo", "2");
ruby_test!(test_module_inclusion_multiple, "module M1; def foo; 1; end; end; module M2; def foo; 2; end; end; class C; include M1; include M2; end; puts C.new.foo", "2");
ruby_test!(test_module_inclusion_ancestors, "module M; end; class C; include M; end; puts C.ancestors.include?(M)", "true");
ruby_test!(test_module_inclusion_included_modules, "module M; end; class C; include M; end; puts C.included_modules.include?(M)", "true");
ruby_test!(test_module_inclusion_extend, "module M; def foo; 1; end; end; class C; extend M; end; puts C.foo", "1");
ruby_test!(test_module_inclusion_extend_object, "module M; def foo; 1; end; end; obj = Object.new; obj.extend(M); puts obj.foo", "1");
ruby_test!(test_module_inclusion_module_function, "module M; module_function; def foo; 1; end; end; puts M.foo", "1");
ruby_test!(test_module_inclusion_module_function_private, "module M; module_function; def foo; 1; end; end; class C; include M; def bar; foo; end; end; puts C.new.bar", "1");
