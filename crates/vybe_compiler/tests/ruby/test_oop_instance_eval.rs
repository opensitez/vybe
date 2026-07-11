
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_instance_eval_basic, "class A; def initialize; @x = 1; end; end; puts A.new.instance_eval { @x }", "1");
ruby_test!(test_instance_eval_string, "class A; def initialize; @x = 1; end; end; puts A.new.instance_eval(\"@x\")", "1");
ruby_test!(test_instance_eval_define_method, "obj = Object.new; obj.instance_eval { def foo; 'foo'; end }; puts obj.foo", "foo"); // defines singleton method
ruby_test!(test_instance_eval_self, "obj = Object.new; puts obj.instance_eval { self } == obj", "true");
ruby_test!(test_instance_exec_basic, "class A; def initialize; @x = 1; end; end; puts A.new.instance_exec(2) {|y| @x + y }", "3");
ruby_test!(test_class_exec_alias, "class A; end; A.class_exec(2) {|x| def foo; 2; end }; puts A.new.foo", "2");
