use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_kernel_printing_puts, "puts 'hello'", "hello");
ruby_test!(test_kernel_printing_print, "print 'hello'", "hello");
ruby_test!(test_kernel_printing_p, "p 'hello'", "\"hello\"");
ruby_test!(test_kernel_printing_sprintf, "puts sprintf('%03d', 5)", "005");
ruby_test!(test_kernel_printing_format, "puts format('%03d', 5)", "005");
ruby_test!(test_kernel_printing_warn, "warn 'hello' 2>/dev/null; puts 'ok'", "ok");
ruby_test!(test_kernel_printing_pp, "require 'pp'; pp 'hello'", "\"hello\"");
