use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_attr_writer_basic, "class A; attr_writer :x; def x; @x; end; end; a = A.new; a.x = 1; puts a.x", "1");
ruby_test!(test_attr_writer_multiple, "class A; attr_writer :x, :y; def xy; \"#{@x}-#{@y}\"; end; end; a = A.new; a.x = 1; a.y = 2; puts a.xy", "1-2");
ruby_test!(test_attr_writer_string, "class A; attr_writer 'x'; def x; @x; end; end; a = A.new; a.x = 1; puts a.x", "1");
ruby_test!(test_attr_writer_missing_method_error, "class A; attr_writer :x; end; a = A.new; a.x = 1; begin; a.x; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_attr_accessor_basic, "class A; attr_accessor :x; end; a = A.new; a.x = 1; puts a.x", "1");
