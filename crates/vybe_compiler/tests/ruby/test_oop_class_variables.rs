macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_class_variable_basic,
    "class A; @@c = 'c'; def foo; @@c; end; end; puts A.new.foo",
    "c"
);
ruby_test!(
    test_class_variable_inheritance,
    "class A; @@c = 'c'; end; class B < A; def foo; @@c; end; end; puts B.new.foo",
    "c"
);
ruby_test!(
    test_class_variable_shared,
    "class A; @@c = 'a'; end; class B < A; @@c = 'b'; end; class A; def foo; @@c; end; end; puts A.new.foo",
    "b"
); // shared across hierarchy
ruby_test!(
    test_class_variable_get,
    "class A; @@c = 'c'; end; puts A.class_variable_get(:@@c)",
    "c"
);
ruby_test!(
    test_class_variable_set,
    "class A; @@c = 'c'; end; A.class_variable_set(:@@c, 'd'); puts A.class_variable_get(:@@c)",
    "d"
);
ruby_test!(
    test_class_variable_defined,
    "class A; @@c = 'c'; end; puts A.class_variable_defined?(:@@c)",
    "true"
);
ruby_test!(
    test_class_variables_list,
    "class A; @@c = 'c'; @@d = 'd'; end; puts A.class_variables.sort.join('-')",
    "@@c-@@d"
);
