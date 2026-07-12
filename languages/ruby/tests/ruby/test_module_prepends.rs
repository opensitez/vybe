macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_module_prepends_basic,
    "module M; def foo; 1; end; end; class C; prepend M; end; puts C.new.foo",
    "1"
);
ruby_test!(
    test_module_prepends_override,
    "module M; def foo; 1; end; end; class C; prepend M; def foo; 2; end; end; puts C.new.foo",
    "1"
); // prepend takes precedence
ruby_test!(
    test_module_prepends_super,
    "module M; def foo; super + 1; end; end; class C; prepend M; def foo; 1; end; end; puts C.new.foo",
    "2"
);
ruby_test!(
    test_module_prepends_multiple,
    "module M1; def foo; super + 1; end; end; module M2; def foo; super + 2; end; end; class C; def foo; 0; end; prepend M1; prepend M2; end; puts C.new.foo",
    "3"
);
ruby_test!(
    test_module_prepends_ancestors,
    "module M; end; class C; prepend M; end; puts C.ancestors.first.name",
    "M"
);
ruby_test!(
    test_module_prepends_included_modules,
    "module M; end; class C; prepend M; end; puts C.included_modules.include?(M)",
    "true"
);
ruby_test!(
    test_module_prepends_module_prepend,
    "module M1; def foo; 1; end; end; module M2; prepend M1; end; class C; include M2; end; puts C.new.foo",
    "1"
);
