
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_exception_custom_class, "class MyError < StandardError; end; begin; raise MyError, 'err'; rescue MyError => e; puts 'caught'; end", "caught");
ruby_test!(test_exception_custom_message, "class MyError < StandardError; def message; 'custom'; end; end; begin; raise MyError; rescue => e; puts e.message; end", "custom");
ruby_test!(test_exception_custom_raise_instance, "class MyError < StandardError; end; begin; raise MyError.new('err'); rescue => e; puts e.message; end", "err");
ruby_test!(test_exception_custom_raise_class, "class MyError < StandardError; end; begin; raise MyError; rescue => e; puts e.class.name; end", "MyError");
ruby_test!(test_exception_custom_raise_class_message, "class MyError < StandardError; end; begin; raise MyError, 'err'; rescue => e; puts \"#{e.class.name}-#{e.message}\"; end", "MyError-err");
ruby_test!(test_exception_custom_rescue_superclass, "class MyError < StandardError; end; begin; raise MyError; rescue StandardError; puts 'caught'; end", "caught");
ruby_test!(test_exception_custom_rescue_module, "module MyModule; end; class MyError < StandardError; include MyModule; end; begin; raise MyError; rescue MyModule; puts 'caught'; end", "caught");
