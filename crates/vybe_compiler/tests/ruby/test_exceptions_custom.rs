use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_custom_exception_basic, "class MyError < StandardError; end; begin; raise MyError; rescue MyError => e; puts e.class.name; end", "MyError");
ruby_test!(test_custom_exception_message, "class MyError < StandardError; end; begin; raise MyError, 'err'; rescue MyError => e; puts e.message; end", "err");
ruby_test!(test_custom_exception_initialize, "class MyError < StandardError; def initialize(msg, code); super(msg); @code = code; end; attr_reader :code; end; begin; raise MyError.new('err', 404); rescue MyError => e; puts \"#{e.message}-#{e.code}\"; end", "err-404");
ruby_test!(test_custom_exception_hierarchy, "class BaseError < StandardError; end; class SubError < BaseError; end; begin; raise SubError; rescue BaseError; puts 'caught'; end", "caught");
