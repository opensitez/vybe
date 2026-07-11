
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_raise_basic, "begin; raise 'err'; rescue => e; puts e.message; end", "err");
ruby_test!(test_raise_no_args, "begin; raise; rescue => e; puts e.class.name; end", "RuntimeError"); // raise without args raises RuntimeError or current exception
ruby_test!(test_raise_class, "begin; raise StandardError; rescue => e; puts e.class.name; end", "StandardError");
ruby_test!(test_raise_class_and_message, "begin; raise ArgumentError, 'bad arg'; rescue => e; puts \"#{e.class.name}-#{e.message}\"; end", "ArgumentError-bad arg");
ruby_test!(test_raise_re_raise, "begin; begin; raise 'err1'; rescue; raise; end; rescue => e; puts e.message; end", "err1"); // raise inside rescue re-raises
ruby_test!(test_fail_alias, "begin; fail 'err'; rescue => e; puts e.message; end", "err"); // fail is alias for raise
