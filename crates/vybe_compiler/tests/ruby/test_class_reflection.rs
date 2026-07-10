use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_class_new, "C = Class.new; puts C.new.class.name", "C");
ruby_test!(test_class_new_superclass, "A = Class.new; B = Class.new(A); puts B.superclass == A", "true");
ruby_test!(test_class_new_block, "C = Class.new { def foo; 1; end }; puts C.new.foo", "1");
ruby_test!(test_class_allocate, "class C; def initialize; @a = 1; end; def foo; @a; end; end; obj = C.allocate; puts obj.foo.nil?", "true");
ruby_test!(test_class_class_variables, "class C; @@a = 1; @@b = 2; end; puts C.class_variables.sort.join('-')", "@@a-@@b");
ruby_test!(test_class_class_variable_get, "class C; @@a = 1; end; puts C.class_variable_get(:@@a)", "1");
ruby_test!(test_class_class_variable_set, "class C; end; C.class_variable_set(:@@a, 1); puts C.class_variable_get(:@@a)", "1");
ruby_test!(test_class_class_variable_defined, "class C; @@a = 1; end; puts C.class_variable_defined?(:@@a)", "true");
ruby_test!(test_class_remove_class_variable, "class C; @@a = 1; end; C.send(:remove_class_variable, :@@a); puts C.class_variable_defined?(:@@a)", "false");
ruby_test!(test_class_name, "class C; end; puts C.name", "C");
ruby_test!(test_class_name_anonymous, "C = Class.new; puts C.name", "C");
ruby_test!(test_class_name_unassigned, "puts Class.new.name.nil?", "true");
