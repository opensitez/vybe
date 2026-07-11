
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_remove_const_basic, "class A; C = 'C'; end; A.send(:remove_const, :C); puts A.const_defined?(:C)", "false");
ruby_test!(test_remove_const_returns_value, "class A; C = 'C'; end; puts A.send(:remove_const, :C)", "C");
ruby_test!(test_remove_const_missing, "class A; end; begin; A.send(:remove_const, :C); rescue NameError; puts 'err'; end", "err");
ruby_test!(test_remove_class_variable_basic, "class A; @@c = 'c'; end; A.send(:remove_class_variable, :@@c); puts A.class_variable_defined?(:@@c)", "false");
ruby_test!(test_remove_class_variable_returns_value, "class A; @@c = 'c'; end; puts A.send(:remove_class_variable, :@@c)", "c");
ruby_test!(test_remove_instance_variable_basic, "class A; def initialize; @x = 1; end; end; a = A.new; a.send(:remove_instance_variable, :@x); puts a.instance_variable_defined?(:@x)", "false");
ruby_test!(test_remove_instance_variable_returns_value, "class A; def initialize; @x = 1; end; end; a = A.new; puts a.send(:remove_instance_variable, :@x)", "1");
