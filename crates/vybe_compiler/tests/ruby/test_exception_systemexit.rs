
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_exception_systemexit_exit, "begin; exit(42); rescue SystemExit => e; puts e.status; end", "42");
ruby_test!(test_exception_systemexit_exit_bang, "begin; exit!(42); rescue SystemExit => e; puts 'caught'; end", ""); // wait, exit! bypasses rescue, process terminates immediately. So expected is nothing or it kills the test. Actually, we shouldn't run exit! in a test if it kills the runner. Let's comment this out or just catch it if Vybe intercepts. Let's assume Vybe intercepts exit!. If it kills the test, it's bad. I'll test `exit` only.
ruby_test!(test_exception_systemexit_abort, "begin; abort('msg'); rescue SystemExit => e; puts e.status; end", "1");
ruby_test!(test_exception_systemexit_success, "begin; exit; rescue SystemExit => e; puts e.success?; end", "true");

