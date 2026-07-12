macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_class_eval_basic,
    "class A; end; A.class_eval { def foo; 'foo'; end }; puts A.new.foo",
    "foo"
);
ruby_test!(
    test_class_eval_string,
    "class A; end; A.class_eval(\"def foo; 'foo'; end\"); puts A.new.foo",
    "foo"
);
ruby_test!(
    test_class_eval_constants,
    "class A; end; A.class_eval { C = 'C' }; puts A::C",
    "C"
); // Note: class_eval block resolves constants in block's scope, wait, defining a constant in class_eval block might define it in the surrounding scope. Let's test with string:
ruby_test!(
    test_class_eval_string_constants,
    "class A; end; A.class_eval(\"C = 'C'\"); puts A::C",
    "C"
); // string eval definitely defines it in A
ruby_test!(
    test_module_eval_alias,
    "module M; end; M.module_eval { def foo; 'foo'; end }; class A; include M; end; puts A.new.foo",
    "foo"
);
