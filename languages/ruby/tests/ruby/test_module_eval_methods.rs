macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_module_module_eval,
    "module M; end; M.module_eval('def foo; 1; end'); class C; include M; end; puts C.new.foo",
    "1"
);
ruby_test!(
    test_module_class_eval,
    "class C; end; C.class_eval('def foo; 1; end'); puts C.new.foo",
    "1"
);
ruby_test!(
    test_module_module_exec,
    "module M; end; M.module_exec(42) { |x| def foo(y); y + 1; end }; class C; include M; end; puts C.new.foo(41)",
    "42"
);
ruby_test!(
    test_module_class_exec,
    "class C; end; C.class_exec(42) { |x| def foo(y); y + 1; end }; puts C.new.foo(41)",
    "42"
);
ruby_test!(
    test_module_const_get,
    "module M; A = 1; end; puts M.const_get(:A)",
    "1"
);
ruby_test!(
    test_module_const_set,
    "module M; end; M.const_set(:A, 1); puts M::A",
    "1"
);
ruby_test!(
    test_module_const_defined,
    "module M; A = 1; end; puts M.const_defined?(:A)",
    "true"
);
ruby_test!(
    test_module_remove_const,
    "module M; A = 1; remove_const(:A); end; puts M.const_defined?(:A)",
    "false"
);
ruby_test!(
    test_module_constants,
    "module M; A = 1; B = 2; end; puts M.constants.sort.join('-')",
    "A-B"
);
ruby_test!(
    test_module_define_method,
    "class C; define_method(:foo) { 1 }; end; puts C.new.foo",
    "1"
);
ruby_test!(
    test_module_alias_method,
    "class C; def foo; 1; end; alias_method :bar, :foo; end; puts C.new.bar",
    "1"
);
ruby_test!(
    test_module_undef_method,
    "class C; def foo; 1; end; undef_method :foo; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_module_remove_method,
    "class A; def foo; 1; end; end; class B < A; def foo; 2; end; remove_method :foo; end; puts B.new.foo",
    "1"
);
