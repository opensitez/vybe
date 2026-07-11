
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_kernel_eval_file_line, "eval('puts [__FILE__, __LINE__].join(\"-\")', nil, 'foo.rb', 42)", "foo.rb-42");
ruby_test!(test_kernel_eval_binding, "a = 1; eval('puts a', binding)", "1");
ruby_test!(test_kernel_local_variables, "a = 1; puts local_variables.include?(:a)", "true");
ruby_test!(test_kernel_iterator, "def foo; puts block_given?; end; foo", "false");
ruby_test!(test_kernel_iterator_true, "def foo; puts block_given?; end; foo {}", "true");
ruby_test!(test_kernel___dir__, "puts __dir__.class.name", "String");
ruby_test!(test_kernel___callee__, "def foo; puts __callee__; end; foo", "foo");
ruby_test!(test_kernel___method__, "def foo; puts __method__; end; foo", "foo");
ruby_test!(test_kernel_loop, "acc = 0; loop do; acc += 1; break if acc == 3; end; puts acc", "3");
