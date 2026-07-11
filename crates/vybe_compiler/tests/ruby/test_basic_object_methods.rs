
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_basic_object_eq, "puts BasicObject.new == BasicObject.new", "false");
ruby_test!(test_basic_object_not_eq, "puts BasicObject.new != BasicObject.new", "true");
ruby_test!(test_basic_object_not, "puts (!BasicObject.new)", "false");
ruby_test!(test_basic_object_instance_eval, "puts BasicObject.new.instance_eval { 42 }", "42");
ruby_test!(test_basic_object_instance_exec, "puts BasicObject.new.instance_exec(42) { |x| x }", "42");
ruby_test!(test_basic_object_singleton_method_added, "class BO < BasicObject; def singleton_method_added(id); end; end; puts BO.new.class.name rescue 'NoMethodError'", "NoMethodError");
