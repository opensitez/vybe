
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_module_attr_accessor, "class C; attr_accessor :x; end; c = C.new; c.x = 1; puts c.x", "1");
ruby_test!(test_module_attr_reader, "class C; attr_reader :x; def init; @x = 1; end; end; c = C.new; c.init; puts c.x", "1");
ruby_test!(test_module_attr_writer, "class C; attr_writer :x; def get_x; @x; end; end; c = C.new; c.x = 1; puts c.get_x", "1");
ruby_test!(test_module_attr, "class C; attr :x, true; end; c = C.new; c.x = 1; puts c.x", "1");
ruby_test!(test_module_attr_false, "class C; attr :x; def init; @x = 1; end; end; c = C.new; c.init; puts c.x", "1");
ruby_test!(test_module_define_method, "class C; define_method(:foo) { 42 }; end; puts C.new.foo", "42");
ruby_test!(test_module_define_method_args, "class C; define_method(:foo) { |a| a * 2 }; end; puts C.new.foo(21)", "42");
ruby_test!(test_module_undef_method, "class C; def foo; 1; end; undef_method :foo; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_module_remove_method, "class C; def foo; 1; end; end; class D < C; def foo; 2; end; remove_method :foo; end; puts D.new.foo", "1");
