use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_kernel_callcc, "puts callcc { |c| c.call(42); 100 }", "42");
ruby_test!(test_kernel_catch_throw, "puts catch(:done) { throw :done, 42; 100 }", "42");
ruby_test!(test_kernel_catch_throw_no_arg, "puts (catch(:done) { throw :done; 100 }).nil?", "true");
ruby_test!(test_kernel_catch_no_throw, "puts catch(:done) { 100 }", "100");
ruby_test!(test_kernel_caller, "def a; b; end; def b; caller; end; puts a[0].include?('b')", "true");
ruby_test!(test_kernel_caller_locations, "def a; b; end; def b; caller_locations; end; puts a[0].class.name", "Thread::Backtrace::Location");
ruby_test!(test_kernel_block_given, "def foo; block_given?; end; puts foo", "false");
ruby_test!(test_kernel_block_given_true, "def foo; block_given?; end; puts foo {}", "true");
ruby_test!(test_kernel_local_variables, "a = 1; b = 2; puts local_variables.sort.join('-')", "a-b");
ruby_test!(test_kernel_global_variables, "puts global_variables.include?(:$!).to_s", "true");
ruby_test!(test_kernel_warn, "warn('test warning'); puts 'done'", "done");
ruby_test!(test_kernel_sleep, "puts sleep(0).class.name", "Integer");
