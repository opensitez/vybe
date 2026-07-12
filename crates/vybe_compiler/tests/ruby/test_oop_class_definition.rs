macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_class_definition_basic,
    "class A; end; puts A.class.name",
    "Class"
);
ruby_test!(
    test_class_definition_methods,
    "class A; def foo; 'foo'; end; end; puts A.new.foo",
    "foo"
);
ruby_test!(
    test_class_reopening,
    "class A; def foo; 'foo'; end; end; class A; def bar; 'bar'; end; end; puts \"#{A.new.foo}-#{A.new.bar}\"",
    "foo-bar"
);
ruby_test!(test_class_name, "class A; end; puts A.name", "A");
ruby_test!(
    test_class_new_block,
    "c = Class.new { def foo; 'foo'; end }; puts c.new.foo",
    "foo"
);
ruby_test!(
    test_class_new_superclass,
    "class A; def foo; 'foo'; end; end; c = Class.new(A); puts c.new.foo",
    "foo"
);
