macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_system_exit_basic,
    "begin; exit; rescue SystemExit => e; puts e.status; end",
    "0"
);
ruby_test!(
    test_system_exit_status,
    "begin; exit(99); rescue SystemExit => e; puts e.status; end",
    "99"
);
ruby_test!(
    test_system_exit_success,
    "begin; exit; rescue SystemExit => e; puts e.success?; end",
    "true"
);
ruby_test!(
    test_system_exit_success_false,
    "begin; exit(1); rescue SystemExit => e; puts e.success?; end",
    "false"
);
ruby_test!(
    test_system_exit_abort,
    "begin; abort 'msg'; rescue SystemExit => e; puts \"#{e.status}-#{e.message}\"; end",
    "1-msg"
);
ruby_test!(
    test_system_exit_not_caught_by_standard_error,
    "begin; exit; rescue StandardError; puts 'caught'; rescue SystemExit; puts 'system_exit'; end",
    "system_exit"
);
