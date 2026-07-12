macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_class_inheritance_basic,
    "class A; def foo; 1; end; end; class B < A; end; puts B.new.foo",
    "1"
);
ruby_test!(
    test_class_inheritance_super,
    "class A; def foo; 1; end; end; class B < A; def foo; super + 1; end; end; puts B.new.foo",
    "2"
);
ruby_test!(
    test_class_inheritance_super_args,
    "class A; def foo(x); x; end; end; class B < A; def foo(x); super(x + 1); end; end; puts B.new.foo(1)",
    "2"
);
ruby_test!(
    test_class_inheritance_super_implicit_args,
    "class A; def foo(x); x; end; end; class B < A; def foo(x); super; end; end; puts B.new.foo(42)",
    "42"
);
ruby_test!(
    test_class_inheritance_class_methods,
    "class A; def self.foo; 1; end; end; class B < A; end; puts B.foo",
    "1"
);
ruby_test!(
    test_class_inheritance_superclass,
    "class A; end; class B < A; end; puts B.superclass.name",
    "A"
);
ruby_test!(
    test_class_inheritance_ancestors,
    "class A; end; class B < A; end; puts B.ancestors.include?(A)",
    "true"
);
ruby_test!(
    test_class_inheritance_object,
    "class A; end; puts A.superclass.name",
    "Object"
);
ruby_test!(
    test_class_inheritance_basicobject,
    "class A < BasicObject; end; puts A.superclass.name",
    "BasicObject"
);
ruby_test!(
    test_class_inheritance_override,
    "class A; def foo; 1; end; end; class B < A; def foo; 2; end; end; puts B.new.foo",
    "2"
);
