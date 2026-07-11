
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_instance_variable_basic, "class A; def initialize; @i = 'i'; end; def foo; @i; end; end; puts A.new.foo", "i");
ruby_test!(test_instance_variable_uninitialized, "class A; def foo; @i; end; end; puts A.new.foo.nil?", "true");
ruby_test!(test_instance_variable_get, "class A; def initialize; @i = 'i'; end; end; puts A.new.instance_variable_get(:@i)", "i");
ruby_test!(test_instance_variable_set, "class A; end; a = A.new; a.instance_variable_set(:@i, 'i'); puts a.instance_variable_get(:@i)", "i");
ruby_test!(test_instance_variable_defined, "class A; def initialize; @i = 'i'; end; end; puts A.new.instance_variable_defined?(:@i)", "true");
ruby_test!(test_instance_variables_list, "class A; def initialize; @i = 'i'; @j = 'j'; end; end; puts A.new.instance_variables.sort.join('-')", "@i-@j");
ruby_test!(test_class_instance_variable, "class A; @i = 'ci'; def self.foo; @i; end; end; puts A.foo", "ci"); // Class instance variable
