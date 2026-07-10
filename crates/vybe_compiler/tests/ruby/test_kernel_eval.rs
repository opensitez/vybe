use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_kernel_eval_basic, "puts eval('1 + 2')", "3");
ruby_test!(test_kernel_eval_binding, "a = 42; puts eval('a', binding)", "42");
ruby_test!(test_kernel_eval_file_line, "begin; eval('raise', nil, 'test.rb', 10); rescue => e; puts e.backtrace.first.include?('test.rb:10'); end", "true");
ruby_test!(test_kernel_eval_local_variables, "a = 1; puts local_variables.include?(:a)", "true");
ruby_test!(test_kernel_eval_global_variables, "puts global_variables.include?(:$0)", "true");
ruby_test!(test_kernel_eval_binding_method, "a = 42; puts binding.local_variable_get(:a)", "42");
ruby_test!(test_kernel_eval_caller, "def foo; caller; end; puts foo.class.name", "Array");
ruby_test!(test_kernel_eval_catch_throw, "puts catch(:done) { throw :done, 42; 0 }", "42");
ruby_test!(test_kernel_eval_loop, "acc = 0; loop { acc += 1; break if acc > 2 }; puts acc", "3");
ruby_test!(test_kernel_eval_block_given, "def foo; puts block_given?; end; foo; foo { }", "false\\ntrue");
