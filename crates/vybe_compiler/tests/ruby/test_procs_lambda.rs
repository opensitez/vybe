use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_lambda_basic, "l = lambda { 'foo' }; puts l.call", "foo");
ruby_test!(test_lambda_stabby_syntax, "l = -> { 'foo' }; puts l.call", "foo");
ruby_test!(test_lambda_stabby_args, "l = ->(x) { \"foo_#{x}\" }; puts l.call(1)", "foo_1");
ruby_test!(test_lambda_predicate_true, "l = lambda { }; puts l.lambda?", "true");
ruby_test!(test_lambda_arity_strict, "l = lambda { |x| x }; begin; l.call(1, 2); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_lambda_return, "def foo; l = lambda { return 'lambda' }; l.call; 'method'; end; puts foo", "method"); // lambda return returns from lambda
ruby_test!(test_proc_return, "def foo; p = Proc.new { return 'proc' }; p.call; 'method'; end; puts foo", "proc"); // proc return returns from enclosing method
