use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_class_new_basic, "C = Class.new; C.class_eval { def foo; 'foo'; end }; puts C.new.foo", "foo");
ruby_test!(test_class_new_superclass, "class A; def foo; 'foo'; end; end; C = Class.new(A); puts C.new.foo", "foo");
ruby_test!(test_class_new_block, "C = Class.new do def foo; 'foo'; end; end; puts C.new.foo", "foo");
ruby_test!(test_module_new_basic, "M = Module.new; M.module_eval { def foo; 'foo'; end }; class A; include M; end; puts A.new.foo", "foo");
ruby_test!(test_module_new_block, "M = Module.new do def foo; 'foo'; end; end; class A; include M; end; puts A.new.foo", "foo");
ruby_test!(test_anonymous_class_name, "c = Class.new; puts c.name.nil?", "true"); // anonymous classes have nil name until assigned to a constant
ruby_test!(test_anonymous_class_assigned_name, "C = Class.new; puts C.name", "C"); // gets named when assigned to constant
