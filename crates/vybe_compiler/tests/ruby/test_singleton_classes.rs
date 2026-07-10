use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_singleton_classes_basic, "obj = Object.new; class << obj; def foo; 1; end; end; puts obj.foo", "1");
ruby_test!(test_singleton_classes_singleton_class_method, "obj = Object.new; puts obj.singleton_class.class.name", "Class");
ruby_test!(test_singleton_classes_define_method, "obj = Object.new; obj.define_singleton_method(:foo) { 1 }; puts obj.foo", "1");
ruby_test!(test_singleton_classes_class_methods, "class C; class << self; def foo; 1; end; end; end; puts C.foo", "1");
ruby_test!(test_singleton_classes_inheritance, "class A; class << self; def foo; 1; end; end; end; class B < A; end; puts B.foo", "1");
ruby_test!(test_singleton_classes_super, "class A; class << self; def foo; 1; end; end; end; class B < A; class << self; def foo; super + 1; end; end; end; puts B.foo", "2");
ruby_test!(test_singleton_classes_module_extend, "module M; def foo; 1; end; end; class C; class << self; include M; end; end; puts C.foo", "1");
ruby_test!(test_singleton_classes_not_instantiable, "obj = Object.new; begin; obj.singleton_class.new; rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_singleton_classes_ancestors, "obj = Object.new; puts obj.singleton_class.ancestors.include?(Object)", "true");
