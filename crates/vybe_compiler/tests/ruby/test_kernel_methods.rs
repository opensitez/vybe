use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_kernel_puts, "puts 'hello'", "hello");
ruby_test!(test_kernel_print, "print 'hello\\n'", "hello");
ruby_test!(test_kernel_p, "p 'hello'", "\"hello\"");
ruby_test!(test_kernel_sprintf, "puts sprintf('%03d', 5)", "005");
ruby_test!(test_kernel_format, "puts format('%03d', 5)", "005");
ruby_test!(test_kernel_warn, "warn 'warning'; puts 'ok'", "ok"); // warn prints to stderr
ruby_test!(test_kernel_raise, "begin; raise 'err'; rescue => e; puts e.message; end", "err");
ruby_test!(test_kernel_loop, "i = 0; loop do i += 1; break if i == 3; end; puts i", "3");
ruby_test!(test_kernel_block_given, "def foo; block_given?; end; puts foo {}", "true");
ruby_test!(test_kernel___dir__, "puts __dir__.is_a?(String) || __dir__.nil?", "true"); // __dir__ is nil in eval/string without path
