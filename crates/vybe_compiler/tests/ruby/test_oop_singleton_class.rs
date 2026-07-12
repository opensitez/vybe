macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_singleton_class_basic,
    "obj = Object.new; class << obj; def foo; 'foo'; end; end; puts obj.foo",
    "foo"
);
ruby_test!(
    test_singleton_method_def,
    "obj = Object.new; def obj.foo; 'foo'; end; puts obj.foo",
    "foo"
);
ruby_test!(
    test_singleton_class_of_class,
    "class A; class << self; def foo; 'foo'; end; end; end; puts A.foo",
    "foo"
);
ruby_test!(
    test_singleton_class_singleton_class,
    "class A; end; puts A.singleton_class.is_a?(Class)",
    "true"
);
ruby_test!(
    test_singleton_class_ancestors,
    "obj = Object.new; puts obj.singleton_class.ancestors.include?(Object)",
    "true"
);
ruby_test!(
    test_singleton_class_inheritance,
    "class A; def self.foo; 'A'; end; end; class B < A; end; puts B.foo",
    "A"
); // Singleton class inherits from superclass's singleton class
